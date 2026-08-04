import streamlit as st
import pandas as pd
import anthropic
import json
import os
import re
from datetime import datetime

# ─────────────────────────────────────────
# 사업장별 키워드 매핑
# ─────────────────────────────────────────
SITE_KEYWORDS = {
    "SR6": {"salary": ["삼성전자기흥", "SR6"], "severance": ["산업체"]},
    "동탄에듀": {"salary": ["동탄", "에듀"], "severance": ["연수원"]},
    "넥슨": {"salary": ["넥슨"], "severance": ["오피스"]}
}

def extract_month_from_filename(filename):
    """파일명에서 월 추출 (예: 2026.04 → 2026년 04월)"""
    match = re.search(r'(\d{4})[.\-_](\d{2})', str(filename))
    if match:
        return f"{match.group(1)}년 {match.group(2)}월"
    return None

def filter_site_data(df, site_name):
    if df is None or df.empty:
        return pd.DataFrame(), pd.DataFrame()
    keywords = SITE_KEYWORDS.get(site_name, {})
    site_col = df.columns[1]
    salary_df = df[df[site_col].apply(lambda x: any(kw in str(x) for kw in keywords.get("salary", [])) if pd.notna(x) else False)]
    sev_df = df[df[site_col].apply(lambda x: any(kw in str(x) for kw in keywords.get("severance", [])) if pd.notna(x) else False)]
    return salary_df, sev_df

def read_excel_sheets(file):
    import io, os
    os.environ["PYTHONIOENCODING"] = "utf-8"
    engine = "xlrd" if str(file.name).endswith(".xls") else "openpyxl"
    file_bytes = file.read()
    xl = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
    sheets = {}
    for sheet in xl.sheet_names:
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, engine=engine, header=4)
            df.columns = [str(c).encode("utf-8", errors="replace").decode("utf-8") for c in df.columns]
            # 모든 문자열 컬럼 utf-8 처리
            for col in df.select_dtypes(include="object").columns:
                df[col] = df[col].apply(lambda x: str(x).encode("utf-8", errors="replace").decode("utf-8") if pd.notna(x) else x)
            sheets[sheet] = df
        except Exception as e:
            pass
    file.seek(0)
    return sheets

def read_extra_cost_sheet(file, site_name):
    try:
        import io
        engine = "xlrd" if str(file.name).endswith(".xls") else "openpyxl"
        file_bytes = file.read()
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="기타비용", engine=engine, header=0)
        file.seek(0)
        df.columns = [str(c) for c in df.columns]
        site_col = df.columns[1]
        keywords = SITE_KEYWORDS.get(site_name, {}).get("salary", [])
        filtered = df[df[site_col].apply(lambda x: any(kw in str(x) for kw in keywords) if pd.notna(x) else False)]
        if filtered.empty:
            return "해당 사업장 기타비용 없음"
        result = []
        for _, row in filtered.iterrows():
            cols = list(row.index)
            name_col = cols[3] if len(cols) > 3 else None
            amount_col = cols[12] if len(cols) > 12 else None
            reason_col = cols[13] if len(cols) > 13 else None
            name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
            amount = row[amount_col] if amount_col and pd.notna(row[amount_col]) else 0
            reason = str(row[reason_col]).strip() if reason_col and pd.notna(row[reason_col]) else ""
            site = str(row[site_col]).strip() if pd.notna(row[site_col]) else ""
            if name and name not in ["nan", ""]:
                result.append(f"  - [{site}] {name} / 사유: {reason} / 금액: {int(float(amount)):,}원")
            else:
                result.append(f"  - [{site}] 보험료정산 / 금액: {int(float(amount)):,}원")
        return "\n".join(result) if result else "기타비용 없음"
    except Exception as e:
        return f"기타비용 시트 읽기 오류: {str(e)}"

def read_profit_file(file):
    try:
        df = pd.read_excel(file)
        month_cols = [c for c in df.columns if '년' in str(c) and '월' in str(c)]
        if not month_cols:
            return df.to_string(), None, None
        non_zero_months = []
        for c in month_cols:
            try:
                col_sum = pd.to_numeric(df[c], errors='coerce').sum()
                if col_sum != 0:
                    non_zero_months.append(c)
            except:
                pass
        if len(non_zero_months) < 2:
            return df.to_string(), None, None
        prev_month = non_zero_months[-2]
        curr_month = non_zero_months[-1]
        id_cols = [c for c in df.columns if '년' not in str(c) and '월' not in str(c) and 'Unnamed' not in str(c)]
        selected = df[id_cols + [prev_month, curr_month]].copy()
        selected['증감'] = pd.to_numeric(selected[curr_month], errors='coerce') - pd.to_numeric(selected[prev_month], errors='coerce')
        result = f"[전월: {prev_month} / 당월: {curr_month}]\n\n" + selected.to_string()
        return result, prev_month, curr_month
    except Exception as e:
        return f"읽기 오류: {str(e)}", None, None

def read_salary_report(file):
    """급여자료보고서 핵심 데이터 구조적 추출"""
    try:
        engine = "xlrd" if str(file.name).endswith(".xls") else "openpyxl"
        df = pd.read_excel(file, engine=engine, header=None)

        def gv(row, col):
            try:
                v = df.iloc[row, col]
                return "" if str(v) == "nan" else v
            except:
                return ""

        def fmt(v):
            try:
                return f"{int(float(v)):,}원"
            except:
                return str(v)

        r = []
        r.append(f"급여월: {gv(3,17)} / 거래처: {gv(3,7)}")
        r.append("")
        r.append("[매출액]")
        r.append(f"  당월: {fmt(gv(6,8))} / 전월: {fmt(gv(6,10))} / 증감: {fmt(gv(6,13))}")
        if gv(6,15): r.append(f"  사유: {gv(6,15)}")
        r.append("")
        r.append("[상용직 급여 상세]")
        items = [(8,"기본급"),(9,"식대비(비과세)"),(10,"연장수당"),(11,"성과급"),(12,"상여금"),(13,"기타수당 및 급여")]
        for ri, name in items:
            ca, pa, da = gv(ri,8), gv(ri,10), gv(ri,13)
            cc, pc, dc = gv(ri,6), gv(ri,9), gv(ri,11)
            if ca or pa:
                r.append(f"  {name}: 당월 {fmt(ca)}({cc}명) / 전월 {fmt(pa)}({pc}명) / 증감 {fmt(da)}({dc}명)")
        ca2, pa2, da2 = gv(17,8), gv(17,10), gv(17,13)
        r.append(f"  → 상용직 계: 당월 {fmt(ca2)} / 전월 {fmt(pa2)} / 증감 {fmt(da2)}")
        r.append("")
        r.append("[일용직 급여]")
        r.append(f"  기본급: 당월 {fmt(gv(18,8))}({gv(18,6)}명) / 전월 {fmt(gv(18,10))}({gv(18,9)}명) / 증감 {fmt(gv(18,13))}({gv(18,11)}명)")
        r.append("")
        r.append("[급여지급율 - 가지급금 포함]")
        r.append(f"  당월: {gv(23,8)} / 전월: {gv(23,10)} / 변동: {gv(23,13)}")
        if gv(23,15): r.append(f"  사유: {gv(23,15)}")
        r.append("[급여지급율 - 가지급금 제외]")
        r.append(f"  당월: {gv(24,8)} / 전월: {gv(24,10)} / 변동: {gv(24,13)}")
        return "\n".join(r)
    except Exception as e:
        return f"읽기 오류: {str(e)}"

# ─────────────────────────────────────────
# 데이터 저장/로드
# ─────────────────────────────────────────
DATA_FILE = "data.json"

def get_default_data():
    return {
        "settings": {
            "system_title": "사업장 관리 시스템",
            "sidebar_title": "사업장 관리 시스템",
            "font": "기본",
            "bg_color": "#ffffff",
            "accent_color": "#1f77b4"
        },
        "clients": {
            "풀무원": {
                "sites": {
                    "SR6": {"meetings": [], "issues": [], "history": []},
                    "동탄에듀": {"meetings": [], "issues": [], "history": []},
                    "넥슨": {"meetings": [], "issues": [], "history": []}
                }
            },
            "용마로지스": {"sites": {"용마로지스": {"meetings": [], "issues": [], "history": []}}},
            "미리디": {"sites": {"미리디": {"meetings": [], "issues": [], "history": []}}},
            "국세": {"sites": {"국세": {"meetings": [], "issues": [], "history": []}}},
            "한국환경공단": {"sites": {"한국환경공단": {"meetings": [], "issues": [], "history": []}}}
        },
        "checklist": {
            "items": ["근태", "청구서", "매출세금계산서 발행", "매출세금계산서 결재", "공문발송", "품의", "운영비 결재상신", "법인카드 전표", "매입세금계산서 발행", "매입세금계산서 결재", "급여작업", "급여내역서", "급여자료보고서 결재", "급여확정"],
            "site_enabled": {},
            "checked": {}
        }
    }

def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                content_str = f.read().strip()
                if not content_str:
                    raise ValueError("빈 파일")
                loaded = json.loads(content_str)
                # 기본값에 없는 키 보완
                default = get_default_data()
                for key in default:
                    if key not in loaded:
                        loaded[key] = default[key]
                return loaded
        except Exception:
            # 오류 시 기존 파일 백업 후 기본값 반환
            import shutil
            try:
                shutil.copy(DATA_FILE, DATA_FILE + ".bak")
            except:
                pass
            return get_default_data()
    return get_default_data()

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

API_KEY = "여기에_CLAUDE_API키_입력"
data = load_data()
settings = data["settings"]

FONTS = {
    "기본": "sans-serif",
    "나눔고딕": "'Nanum Gothic', sans-serif",
    "나눔명조": "'Nanum Myeongjo', serif",
    "고딕A1": "'Gothic A1', sans-serif",
    "IBM Plex": "'IBM Plex Sans KR', sans-serif",
}

font_css = FONTS.get(settings.get("font", "기본"), "sans-serif")
bg_color = settings.get("bg_color", "#ffffff")
accent = settings.get("accent_color", "#1f77b4")

st.set_page_config(page_title=settings["system_title"], page_icon="🏢", layout="wide")
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nanum+Gothic&family=Nanum+Myeongjo&family=Gothic+A1&family=IBM+Plex+Sans+KR&display=swap');
html, body, [class*="css"] {{ font-family: {font_css} !important; background-color: {bg_color} !important; }}
.card {{ background: white; border-left: 4px solid {accent}; padding: 12px 16px; border-radius: 8px; margin-bottom: 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# 사이드바
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown(f"## 🏢 {settings['sidebar_title']}")
    st.markdown("---")
    menu_items = ["🏠 홈"]
    menu_map = {}
    for client, cdata in data["clients"].items():
        menu_items.append(f"👥 {client}")
        menu_map[f"👥 {client}"] = (client, None)
        for site in cdata["sites"].keys():
            label = f"　🍽️ {site}"
            menu_items.append(label)
            menu_map[label] = (client, site)
    menu_items.append("📋 월초체크리스트")
    menu_items.append("⚙️ 설정")
    menu = st.radio("메뉴", menu_items)

# ─────────────────────────────────────────
# 🏠 홈
# ─────────────────────────────────────────
if menu == "🏠 홈":
    st.title(f"🏢 {settings['system_title']}")
    st.markdown("---")
    for client, cdata in data["clients"].items():
        st.markdown(f"### 👥 {client}")
        cols = st.columns(max(len(cdata["sites"]), 1))
        for i, (site, sdata) in enumerate(cdata["sites"].items()):
            with cols[i]:
                issue_count = len([x for x in sdata["issues"] if x.get("status") == "미처리"])
                st.markdown(f"<div class='card'><b>🍽️ {site}</b><br>📝 회의록: {len(sdata['meetings'])}건<br>⚠️ 미처리 이슈: {issue_count}건<br>📋 히스토리: {len(sdata['history'])}건</div>", unsafe_allow_html=True)
    st.markdown("---")
    st.subheader("⚠️ 전체 미처리 이슈")
    found = False
    for client, cdata in data["clients"].items():
        for site, sdata in cdata["sites"].items():
            for issue in sdata["issues"]:
                if issue.get("status") == "미처리":
                    found = True
                    st.error(f"**[{client} - {site}]** {issue['date']} - {issue['title']}")
    if not found:
        st.success("✅ 현재 미처리 이슈가 없습니다!")

# ─────────────────────────────────────────
# 👥 고객사
# ─────────────────────────────────────────
elif menu in menu_map and menu_map[menu][1] is None:
    client = menu_map[menu][0]
    cdata = data["clients"][client]
    st.title(f"👥 {client}")
    st.markdown("---")
    cols = st.columns(max(len(cdata["sites"]), 1))
    for i, (site, sdata) in enumerate(cdata["sites"].items()):
        with cols[i]:
            issue_count = len([x for x in sdata["issues"] if x.get("status") == "미처리"])
            st.markdown(f"<div class='card'><b>🍽️ {site}</b><br>📝 회의록: {len(sdata['meetings'])}건<br>⚠️ 미처리 이슈: {issue_count}건<br>📋 히스토리: {len(sdata['history'])}건</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────
# 🍽️ 사업장
# ─────────────────────────────────────────
elif menu in menu_map and menu_map[menu][1] is not None:
    client, site = menu_map[menu]
    sdata = data["clients"][client]["sites"][site]
    sk = f"{client}_{site}"

    st.title(f"🍽️ {client} - {site}")
    st.markdown("---")

    tab1, tab2, tab3, tab4 = st.tabs(["📊 AI 분석", "📝 회의록", "⚠️ 이슈사항", "📋 히스토리"])

    with tab1:
        st.markdown("### 📊 AI 분석")
        has_step1 = f"s1_{sk}" in st.session_state
        if has_step1:
            s1_month = st.session_state.get(f"s1_month_{sk}", "")
            st.success(f"✅ Step1 청구서 분석 결과 있음 {s1_month} → Step2/3에서 참고 가능")
            if st.button("🗑️ Step1 결과 초기화", key=f"clear1_{sk}"):
                del st.session_state[f"s1_{sk}"]
                st.rerun()

        step = st.radio("분석 단계", [
            "Step 1 - 청구서 분석",
            "Step 2 - 급여자료보고서 분석",
            "Step 3 - 매출이익 분석"
        ], horizontal=True, key=f"step_{sk}")

        # ── Step 1 ──
        if step == "Step 1 - 청구서 분석":
            st.info(f"📌 {site} 관련 데이터만 자동 필터링됩니다.")
            memo = st.text_area(
                "📝 참고사항 메모 (선택) - AI 분석 시 반영됩니다",
                placeholder="예: 전월 연차수당 과지급건 당월 환급 / 안전인센티브 지급 등",
                key=f"memo_{sk}", height=70
            )
            col1, col2 = st.columns(2)
            with col1:
                prev_file = st.file_uploader("전월 청구서", type=["xlsx", "xls"], key=f"p1_{sk}")
            with col2:
                curr_file = st.file_uploader("당월 청구서", type=["xlsx", "xls"], key=f"c1_{sk}")

            if prev_file and curr_file:
                prev_month = extract_month_from_filename(prev_file.name) or "전월"
                curr_month = extract_month_from_filename(curr_file.name) or "당월"
                st.success(f"✅ {prev_month} / {curr_month} 업로드 완료!")

                if st.button("🤖 청구서 분석 시작", type="primary", key=f"btn1_{sk}"):
                    with st.spinner("분석 중..."):
                        try:
                            prev_sheets = read_excel_sheets(prev_file)
                            curr_sheets = read_excel_sheets(curr_file)
                            prev_extra = read_extra_cost_sheet(prev_file, site)
                            curr_extra = read_extra_cost_sheet(curr_file, site)

                            def get_filtered_text(sheets, label):
                                text = ""
                                for sheet_name, df in sheets.items():
                                    if sheet_name == "기타비용":
                                        continue
                                    salary_df, sev_df = filter_site_data(df, site)
                                    if not salary_df.empty:
                                        try:
                                            s = salary_df.to_string()
                                        except:
                                            s = salary_df.to_string(encoding="utf-8") if hasattr(salary_df.to_string, "__call__") else str(salary_df)
                                        text += f"\n[{label} - {sheet_name} - 급여]\n{s}\n"
                                    if not sev_df.empty:
                                        try:
                                            s = sev_df.to_string()
                                        except:
                                            s = str(sev_df)
                                        text += f"\n[{label} - {sheet_name} - 퇴직금]\n{s}\n"
                                return text

                            prev_text = get_filtered_text(prev_sheets, prev_month)
                            curr_text = get_filtered_text(curr_sheets, curr_month)

                            if not prev_text and not curr_text:
                                st.warning("해당 사업장 데이터를 찾지 못했습니다.")
                            else:
                                client_obj = anthropic.Anthropic(api_key=API_KEY)
                                prompt = f"""당신은 노무/급여 전문가입니다. {site} 청구서(VAT 포함 기준)를 분석해주세요.

분석 기준월: 전월={prev_month}, 당월={curr_month}
모든 분석에서 "전월/당월" 대신 반드시 "{prev_month}/{curr_month}" 실제 월명을 사용하세요.

참고사항:
- 고정비 = 기본급 (시급직은 근무일수에 따라 변동)
- 변동비 = 연장/휴일 근무수당
- 퇴직금: 매달 6.5% 청구, 퇴사 시 법정 8.33%와 차액 익월 청구
  * 1년 이상 퇴사자 → 추가 청구 (퇴직금 증가)
  * 1년 미만 퇴사자 → 기청구분 환급 (퇴직금 감소)
- 기타비용: 이름 있는 항목 → 사유+금액 명시 / 이름 없는 항목 → 보험료 정산
- 인원변동: "중도퇴사 N명, 신규입사 N명" 후 실명 괄호 안에
- 악화/개선 절대 금지. 상승/하락/증가/감소만 사용

[담당자 메모]
{memo if memo else "없음"}

[{prev_month} 급여/퇴직금 데이터]
{prev_text}

[{curr_month} 급여/퇴직금 데이터]
{curr_text}

[{prev_month} 기타비용 상세]
{prev_extra}

[{curr_month} 기타비용 상세]
{curr_extra}

아래 형식으로 작성하세요. 각 항목은 줄을 나눠 간결하게:

## 📋 {site} 청구서 분석 ({curr_month} 기준, VAT 포함)

### 급여 현황
| 항목 | {prev_month} | {curr_month} | 증감 |
|------|-----:|-----:|-----:|
| 총 청구액(매출) | | | |
| 고정비(기본급) | | | |
| 변동비(연장/휴일) | | | |
| 일용직 | | | |
| 기타비용 | | | |

### 퇴직금 (별도)
| 항목 | {prev_month} | {curr_month} | 증감 |
|------|-----:|-----:|-----:|
| 퇴직금 청구액 | | | |

## 🔍 변동 원인
- **매출**: (원인)
- **고정비**: (인원현황 먼저, 실명 괄호)
- **변동비**: (연장/휴일 변동, 주요 인원 괄호)
- **기타비용**: (항목별 한 줄씩)
- **퇴직금**: (금액 패턴으로 원인 추론, 한 줄)
- **특이사항**: (없으면 생략)

숫자 천단위 콤마. 간결하게."""

                                msg = client_obj.messages.create(
                                    model="claude-sonnet-4-6",
                                    max_tokens=2000,
                                    messages=[{"role": "user", "content": prompt}]
                                )
                                st.session_state[f"s1_{sk}"] = msg.content[0].text
                                st.session_state[f"s1_month_{sk}"] = f"({prev_month}→{curr_month})"
                                st.session_state[f"s1_prev_month_{sk}"] = prev_month
                                st.session_state[f"s1_curr_month_{sk}"] = curr_month
                                st.rerun()

                        except Exception as e:
                            st.error(f"오류: {str(e)}")

            if f"s1_{sk}" in st.session_state:
                st.markdown("---")
                st.markdown(st.session_state[f"s1_{sk}"])
                if st.button("📋 히스토리 저장", key=f"sv1_{sk}"):
                    sdata["history"].append({
                        "date": datetime.now().strftime("%Y-%m-%d"),
                        "type": "청구서 분석",
                        "content": st.session_state[f"s1_{sk}"]
                    })
                    save_data(data)
                    st.success("저장됨!")

        # ── Step 2 ──
        elif step == "Step 2 - 급여자료보고서 분석":
            if has_step1:
                st.info("📌 Step1 청구서 분석 결과를 참고합니다.")
            else:
                st.info("📌 Step1 없이 급여자료보고서만으로 분석합니다.")

            sal_file = st.file_uploader("급여자료보고서", type=["xlsx", "xls"], key=f"sal_{sk}")

            if sal_file:
                sal_month = extract_month_from_filename(sal_file.name) or "당월"
                # 전월 추정
                match = re.search(r'(\d{4})년 (\d{2})월', sal_month)
                if match:
                    y, m = int(match.group(1)), int(match.group(2))
                    prev_m = 12 if m == 1 else m - 1
                    prev_y = y - 1 if m == 1 else y
                    sal_prev_month = f"{prev_y}년 {prev_m:02d}월"
                else:
                    sal_prev_month = "전월"

                st.success(f"✅ {sal_month} 급여자료보고서 업로드 완료!")

                if st.button("🤖 급여보고서 분석 시작", type="primary", key=f"btn2_{sk}"):
                    with st.spinner("분석 중..."):
                        try:
                            sal_text = read_salary_report(sal_file)
                            step1_ref = st.session_state.get(f"s1_{sk}", "없음 (청구서 미업로드)")
                            s1_prev = st.session_state.get(f"s1_prev_month_{sk}", sal_prev_month)
                            s1_curr = st.session_state.get(f"s1_curr_month_{sk}", sal_month)

                            client_obj = anthropic.Anthropic(api_key=API_KEY)
                            prompt = f"""당신은 노무/급여 전문가입니다.
{site} {s1_curr} 급여자료보고서를 분석하여, 담당자가 ERP에 입력할 증감사유 멘트를 작성해주세요.

※ 이 멘트는 담당자가 ERP 급여자료보고서의 증감사유 칸에 그대로 입력할 내용입니다.
※ 급여자료보고서 수치(VAT 제외) 기준으로 작성하세요.
※ 청구서는 참고용이며, 금액은 급여자료보고서 기준으로 작성하세요.

분석 기준월: 전월={s1_prev}, 당월={s1_curr}
모든 분석에서 반드시 "{s1_prev}/{s1_curr}" 실제 월명을 사용하세요.

작성 규칙:
- 증감사유 순서: 매출 → 기본급 → 식대 → 연장수당 → 기타수당 → 성과급/상여(있을때만) → 일용직
- 인원변동: "중도퇴사 N명, 신규입사 N명" 후 실명 괄호
- 기타수당 = 연장수당 제외한 모든 수당(휴일/심야/휴일연장 등)
- 변동 없는 항목 생략
- 악화/개선 금지. 상승/하락/증가/감소만 사용
- 각 항목 원인은 줄 나눠서 간결하게

[급여자료보고서 수치]
{sal_text}

[청구서 분석 참고 - 인원/수당 상세 내용]
{step1_ref}

아래 형식으로 작성하세요:

## 📝 {site} {s1_curr} 급여자료보고서

| 항목 | {s1_prev} | {s1_curr} | 증감 |
|------|-----:|-----:|-----:|
| 매출(VAT제외) | | | |
| 상용직 인원 | | | |
| 상용직 소득계 | | | |
| 일용직 인원 | | | |
| 일용직 소득계 | | | |
| 총 소득계 | | | |

## ✏️ ERP 입력용 증감사유

1. 매출: {s1_prev} N원 → {s1_curr} N원 (±N원)
   - (원인. 구체적으로)
2. 기본급: ±N원
   - (원인. 인원변동 상세, 실명 괄호)
3. 식대: ±N원 (변동 있을 때만)
   - (원인)
4. 연장수당: ±N원
   - (원인. 주요 인원 괄호)
5. 기타수당: ±N원 (있을 때만)
   - (원인: 휴일/심야수당 등 구체적으로)
6. 성과급/상여금: (있을 때만)
   - (원인)
7. 일용직: ±N원 (있을 때만)
   - (원인)

## 💡 급여지급율 분석 (★핵심★)

| 구분 | {s1_prev} | {s1_curr} | 변동 |
|------|-----:|-----:|-----:|
| 급여지급율(가지급금 포함) | % | % | %p |
| 급여지급율(가지급금 제외) | % | % | %p |

급여지급율이 변동한 이유:
- (인과관계로 설명. 예: "{s1_prev} 연차수당 과청구분 환급으로 {s1_curr} 매출 N원 감소, 노무비는 유지되어 지급율 N%p 상승")
- (추가 원인 있으면)

## 💬 상사 보고용 한줄 요약
> (핵심 1~2문장. 급여지급율 변동 원인 반드시 포함)

숫자 천단위 콤마. 간결하게."""

                            msg = client_obj.messages.create(
                                model="claude-sonnet-4-6",
                                max_tokens=2000,
                                messages=[{"role": "user", "content": prompt}]
                            )
                            st.session_state[f"s2_{sk}"] = msg.content[0].text

                        except Exception as e:
                            st.error(f"오류: {str(e)}")

            if f"s2_{sk}" in st.session_state:
                st.markdown("---")
                st.markdown(st.session_state[f"s2_{sk}"])
                if st.button("📋 히스토리 저장", key=f"sv2_{sk}"):
                    sdata["history"].append({
                        "date": datetime.now().strftime("%Y-%m-%d"),
                        "type": "급여보고서 분석",
                        "content": st.session_state[f"s2_{sk}"]
                    })
                    save_data(data)
                    st.success("저장됨!")

        # ── Step 3 ──
        elif step == "Step 3 - 매출이익 분석":
            if has_step1:
                st.info("📌 Step1 청구서 분석 결과를 참고합니다.")
            else:
                st.info("📌 Step1 없이 매출이익 파일만으로 분석합니다.")

            profit_file = st.file_uploader("거래처 월별 매출이익", type=["xlsx", "xls"], key=f"pf_{sk}")

            if profit_file:
                st.success("✅ 업로드 완료!")
                if st.button("🤖 매출이익 분석 시작", type="primary", key=f"btn3_{sk}"):
                    with st.spinner("분석 중..."):
                        try:
                            profit_text, prev_month, curr_month = read_profit_file(profit_file)
                            if not prev_month:
                                prev_month = st.session_state.get(f"s1_prev_month_{sk}", "전월")
                                curr_month = st.session_state.get(f"s1_curr_month_{sk}", "당월")
                            step1_ref = st.session_state.get(f"s1_{sk}", "없음 (청구서 미업로드)")

                            client_obj = anthropic.Anthropic(api_key=API_KEY)
                            prompt = f"""당신은 노무/급여 전문가입니다. {site} 매출이익(VAT 제외 기준)을 분석해주세요.

분석 기준월: 전월={prev_month}, 당월={curr_month}
모든 분석에서 "전월/당월" 대신 반드시 "{prev_month}/{curr_month}" 실제 월명을 사용하세요.

중요:
- 거래처 월별 매출이익 파일 수치(VAT 제외)를 기준으로 작성
- 청구서(VAT 포함)와 금액이 다를 수 있으므로 매출이익 파일 수치 우선
- 인원변동: "중도퇴사 N명, 신규입사 N명" 후 실명 괄호
- 퇴직충당금: 크게 증가→장기근속 퇴사자 차액분(8.33%-6.5%) 청구 추정 / 감소→단기근속 퇴사자 환급 또는 인원 감소 추정
- ERP 이슈사항 "가. 원인" 문구 절대 금지. 원인 내용만 바로 작성
- 악화/개선 절대 금지. 상승/하락/증가/감소만 사용
- 각 항목 줄 나눠서 간결하게

[거래처 월별 매출이익 (VAT 제외)]
{profit_text}

[Step1 청구서 분석 참고용]
{step1_ref}

아래 형식으로 작성하세요:

## 📊 {site} 매출이익 현황 ({curr_month})

| 항목 | {prev_month} | {curr_month} | 증감 |
|------|-----:|-----:|-----:|
| 인원(명) | | | |
| 매출(A) | | | |
| 직접원가(B) | | | |
| 직접노무비 | | | |
| 매출이익(A-B) | | | |
| 노무비율 | % | % | %p |
| 매출이익률 | % | % | %p |

## 💡 급여지급율 (노무비/매출액)
| 구분 | {prev_month} | {curr_month} | 변동 |
|------|-----:|-----:|-----:|
| 급여지급율 | % | % | %p |

- {prev_month}: 노무비 ÷ 매출액 = %
- {curr_month}: 노무비 ÷ 매출액 = %
- 변동 원인: (한 줄)

## 📝 ERP 입력용 - 당월 이슈사항

1. 인원: {prev_month} N명 → {curr_month} N명 (±N명)
   - (인원 변동 상세. 실명 괄호)

2. 매출: {prev_month} N원 → {curr_month} N원 (±N원)
   - (원인 1)
   - (원인 2, 있으면)

3. 매출이익: {prev_month} N원 → {curr_month} N원 (±N원)
   - (원인 1)
   - (원인 2, 있으면)

## 💬 상사 보고용 요약
> (핵심 1~2문장. 실제 월명 사용)

숫자 천단위 콤마. 비율 소수점 1자리. 간결하게."""

                            msg = client_obj.messages.create(
                                model="claude-sonnet-4-6",
                                max_tokens=2000,
                                messages=[{"role": "user", "content": prompt}]
                            )
                            st.session_state[f"s3_{sk}"] = msg.content[0].text

                        except Exception as e:
                            st.error(f"오류: {str(e)}")

            if f"s3_{sk}" in st.session_state:
                st.markdown("---")
                st.markdown(st.session_state[f"s3_{sk}"])
                if st.button("📋 히스토리 저장", key=f"sv3_{sk}"):
                    sdata["history"].append({
                        "date": datetime.now().strftime("%Y-%m-%d"),
                        "type": "매출이익 분석",
                        "content": st.session_state[f"s3_{sk}"]
                    })
                    save_data(data)
                    st.success("저장됨!")

            # ── 예상수지 ──
            if profit_file:
                st.markdown("---")
                st.markdown("### 📈 익월 예상수지")
                forecast_memo = st.text_area(
                    "📝 예상수지 참고 메모 (선택)",
                    placeholder="예: 6월 세척실계절수당 약 50만원 지급예정 / 신규입사 2명 예정 등",
                    key=f"fmemo_{sk}", height=70
                )
                if st.button("🔮 예상수지 분석", type="primary", key=f"btn_fc_{sk}"):
                    with st.spinner("예상수지 계산 중..."):
                        try:
                            fc_text, fc_prev, fc_curr = read_profit_file(profit_file)
                            fc_step1 = st.session_state.get(f"s1_{sk}", "없음")

                            client_fc = anthropic.Anthropic(api_key=API_KEY)
                            fc_prompt = f"""당신은 노무/급여 전문가입니다. {site}의 익월 예상수지를 산출해주세요.

데이터: {fc_text}

규칙:
- 최근 3~5개월 데이터 사용
- 상여금이 있는 달(명절 달)은 평균 계산에서 자동 제외
- 제외된 달이 있으면 어떤 달을 제외했는지 명시
- 담당자 메모 내용을 반영하여 예상수지 조정
- 모든 항목(인원/매출/급여/상여/퇴직충당금/잡급/노무비소계/복리후생비/접대비/법정복리후생비/인력수급비/지급수수료/소모품비/통신비/기타비용/경비소계/직접원가계/영업이익)을 채워주세요
- 숫자는 천단위 콤마
- 악화/개선 금지. 상승/하락/증가/감소만 사용

[담당자 메모 - 반드시 반영]
{forecast_memo if forecast_memo else "없음"}

[청구서 참고]
{fc_step1}

아래 형식으로 작성하세요:

## 📈 {site} 익월 예상수지

※ 평균 산출 기준: (사용한 월 명시, 제외한 달과 이유 명시)
※ 메모 반영 내용: (메모 있으면 어떻게 반영했는지)

| 항목 | 평균(N개월) | 익월 예상 | 비고 |
|------|-----:|-----:|------|
| 인원(명) | | | |
| 매출 | | | |
| 급여 | | | |
| 상여 | | 0 | 명절 없음 |
| 퇴직충당금 | | | |
| 잡급(일용직) | | | |
| 노무비 소계 | | | |
| 복리후생비 | | | |
| 접대비 | | | |
| 법정복리후생비 | | | |
| 인력수급비 | | | |
| 지급수수료 | | | |
| 소모품비 | | | |
| 통신비 | | | |
| 기타비용 | | | |
| 경비 소계 | | | |
| 직접원가 계 | | | |
| 영업이익(매출이익) | | | |
| 노무비율 | % | % | |
| 매출이익률 | % | % | |
| 급여지급율 | % | % | |

## 💬 예상수지 요약
> (핵심 내용 1~2문장. 전월 대비 변동 예상 포함)"""

                            fc_msg = client_fc.messages.create(
                                model="claude-sonnet-4-6",
                                max_tokens=2000,
                                messages=[{"role": "user", "content": fc_prompt}]
                            )
                            st.session_state[f"s3_fc_{sk}"] = fc_msg.content[0].text

                        except Exception as e:
                            st.error(f"오류: {str(e)}")

                if f"s3_fc_{sk}" in st.session_state:
                    st.markdown(st.session_state[f"s3_fc_{sk}"])
                    if st.button("📋 히스토리 저장 (예상수지)", key=f"sv_fc_{sk}"):
                        sdata["history"].append({
                            "date": datetime.now().strftime("%Y-%m-%d"),
                            "type": "예상수지",
                            "content": st.session_state[f"s3_fc_{sk}"]
                        })
                        save_data(data)
                        st.success("저장됨!")

    # ── 회의록 ──
    with tab2:
        st.markdown("### 📝 회의록")
        with st.expander("➕ 새 회의록 작성"):
            m_date = st.date_input("날짜", key=f"md_{sk}")
            m_title = st.text_input("제목", key=f"mt_{sk}")
            m_participants = st.text_input("참석자", key=f"mp_{sk}")
            m_content = st.text_area("내용", height=200, key=f"mc_{sk}")
            if st.button("저장", key=f"sm_{sk}"):
                if m_title:
                    sdata["meetings"].append({"date": str(m_date), "title": m_title, "participants": m_participants, "content": m_content})
                    save_data(data)
                    st.success("저장됨!")
                    st.rerun()
        for i, m in enumerate(reversed(sdata["meetings"])):
            with st.expander(f"📝 {m['date']} - {m['title']}"):
                st.write(f"**참석자:** {m.get('participants','')}")
                st.write(m["content"])
                if st.button("삭제", key=f"dm_{sk}_{i}"):
                    sdata["meetings"].pop(len(sdata["meetings"])-1-i)
                    save_data(data)
                    st.rerun()

    # ── 이슈사항 ──
    with tab3:
        st.markdown("### ⚠️ 이슈사항")
        with st.expander("➕ 새 이슈 등록"):
            i_date = st.date_input("날짜", key=f"id_{sk}")
            i_title = st.text_input("제목", key=f"it_{sk}")
            i_content = st.text_area("내용", height=150, key=f"ic_{sk}")
            i_status = st.selectbox("상태", ["미처리", "처리중", "완료"], key=f"is_{sk}")
            if st.button("등록", key=f"si_{sk}"):
                if i_title:
                    sdata["issues"].append({"date": str(i_date), "title": i_title, "content": i_content, "status": i_status})
                    save_data(data)
                    st.success("등록됨!")
                    st.rerun()
        filter_status = st.selectbox("상태 필터", ["전체", "미처리", "처리중", "완료"], key=f"if_{sk}")
        for i, issue in enumerate(reversed(sdata["issues"])):
            if filter_status != "전체" and issue["status"] != filter_status:
                continue
            color = {"미처리": "🔴", "처리중": "🟡", "완료": "🟢"}.get(issue["status"], "⚪")
            with st.expander(f"{color} {issue['date']} - {issue['title']} [{issue['status']}]"):
                st.write(issue["content"])
                new_status = st.selectbox("상태 변경", ["미처리", "처리중", "완료"],
                    index=["미처리", "처리중", "완료"].index(issue["status"]), key=f"uis_{sk}_{i}")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("상태 저장", key=f"uib_{sk}_{i}"):
                        sdata["issues"][len(sdata["issues"])-1-i]["status"] = new_status
                        save_data(data)
                        st.rerun()
                with c2:
                    if st.button("삭제", key=f"dib_{sk}_{i}"):
                        sdata["issues"].pop(len(sdata["issues"])-1-i)
                        save_data(data)
                        st.rerun()

    # ── 히스토리 ──
    with tab4:
        st.markdown("### 📋 AI 분석 히스토리")
        if not sdata["history"]:
            st.info("저장된 분석 결과가 없습니다.")
        for i, h in enumerate(reversed(sdata["history"])):
            with st.expander(f"📄 {h['date']} - {h['type']}"):
                st.markdown(h["content"])
                if st.button("삭제", key=f"dh_{sk}_{i}"):
                    sdata["history"].pop(len(sdata["history"])-1-i)
                    save_data(data)
                    st.rerun()


# ─────────────────────────────────────────
# 📋 월초체크리스트
# ─────────────────────────────────────────
elif menu == "📋 월초체크리스트":
    st.title("📋 월초체크리스트")
    st.markdown("---")

    # 체크리스트 데이터 초기화
    if "checklist" not in data:
        data["checklist"] = {
            "items": ["근태", "청구서", "매출세금계산서 발행", "매출세금계산서 결재",
                     "공문발송", "품의", "운영비 결재상신", "법인카드 전표",
                     "매입세금계산서 발행", "매입세금계산서 결재", "급여작업",
                     "급여내역서", "급여자료보고서 결재", "급여확정"],
            "site_enabled": {},
            "checked": {}
        }

    cl = data["checklist"]
    items = cl.get("items", [])

    # 전체 사업장 목록
    all_sites = []
    for c, cd in data["clients"].items():
        for s in cd["sites"].keys():
            all_sites.append(f"{c} - {s}")

    # site_enabled 초기화 (새 사업장/항목만 추가, 기존 설정 유지)
    for site_key in all_sites:
        if site_key not in cl["site_enabled"]:
            cl["site_enabled"][site_key] = {}
        for item in items:
            if item not in cl["site_enabled"][site_key]:
                cl["site_enabled"][site_key][item] = True

    # checked 초기화
    for site_key in all_sites:
        if site_key not in cl["checked"]:
            cl["checked"][site_key] = {}

    # 현재 월
    curr_month = datetime.now().strftime("%Y년 %m월")

    # 상단 버튼
    col_title, col_btn = st.columns([3, 1])
    with col_title:
        st.subheader(f"📅 {curr_month}")
    with col_btn:
        if st.button("🔄 전체 초기화", type="secondary", key="cl_reset"):
            for site_key in all_sites:
                cl["checked"][site_key] = {}
            save_data(data)
            # 세션 상태 전체 체크박스 키 삭제
            for k in list(st.session_state.keys()):
                if k.startswith("cb_") or k.startswith("cl_"):
                    del st.session_state[k]
            st.success("✅ 전체 초기화 완료!")
            st.rerun()

    st.markdown("---")

    # ── 체크리스트 테이블 ──
    # 헤더
    header_cols = st.columns([2] + [1] * len(all_sites))
    with header_cols[0]:
        st.markdown("**업무항목**")
    for i, site_key in enumerate(all_sites):
        with header_cols[i+1]:
            site_short = site_key.split(" - ")[1] if " - " in site_key else site_key
            st.markdown(f"**{site_short}**")

    st.markdown("---")

    # 항목별 행
    for idx, item in enumerate(items):
        # 홀짝 줄 배경색
        bg = "#f8f9fa" if idx % 2 == 0 else "#ffffff"
        st.markdown(f"""<div style="background:{bg}; padding:4px 8px; border-radius:4px; margin:1px 0;">""", unsafe_allow_html=True)
        row_cols = st.columns([2] + [1] * len(all_sites))
        with row_cols[0]:
            any_checked = any(cl["checked"].get(sk, {}).get(item, False) for sk in all_sites)
            if any_checked:
                st.markdown(f"<span style='color:gray;text-decoration:line-through'>{item}</span>", unsafe_allow_html=True)
            else:
                st.markdown(item)

        for i, site_key in enumerate(all_sites):
            with row_cols[i+1]:
                enabled = cl["site_enabled"].get(site_key, {}).get(item, True)
                if enabled:
                    checked = cl["checked"].get(site_key, {}).get(item, False)
                    # 세션 상태 직접 사용
                    cb_key = f"cb_{site_key}_{item}"
                    if cb_key not in st.session_state:
                        st.session_state[cb_key] = checked
                    new_val = st.checkbox("", key=cb_key, label_visibility="collapsed")
                    if new_val != checked:
                        if site_key not in cl["checked"]:
                            cl["checked"][site_key] = {}
                        cl["checked"][site_key][item] = new_val
                        save_data(data)
                else:
                    st.markdown("—")

    st.markdown("---")

    # ── 항목 관리 ──
    with st.expander("⚙️ 항목 관리"):
        st.markdown("**항목 추가**")
        with st.form("add_item"):
            new_item = st.text_input("새 항목 이름")
            if st.form_submit_button("추가"):
                if new_item and new_item not in cl["items"]:
                    cl["items"].append(new_item)
                    for sk in all_sites:
                        if sk not in cl["site_enabled"]:
                            cl["site_enabled"][sk] = {}
                        cl["site_enabled"][sk][new_item] = True
                    save_data(data)
                    st.success(f"'{new_item}' 추가됨!")
                    st.rerun()

        st.markdown("**항목 삭제**")
        with st.form("del_item"):
            del_item = st.selectbox("삭제할 항목", ["선택안함"] + items)
            if st.form_submit_button("삭제"):
                if del_item != "선택안함":
                    cl["items"].remove(del_item)
                    for sk in all_sites:
                        cl["site_enabled"].get(sk, {}).pop(del_item, None)
                        cl["checked"].get(sk, {}).pop(del_item, None)
                    save_data(data)
                    st.success(f"'{del_item}' 삭제됨!")
                    st.rerun()

        st.markdown("**사업장별 항목 활성화/비활성화**")
        st.caption("체크 해제하면 해당 사업장에서 항목이 '—'로 표시됩니다")
        sel_site = st.selectbox("사업장 선택", all_sites, key="cl_site_sel")
        if sel_site:
            with st.form("toggle_items"):
                toggles = {}
                for item in items:
                    cur = cl["site_enabled"].get(sel_site, {}).get(item, True)
                    toggles[item] = st.checkbox(item, value=cur, key=f"tog_{sel_site}_{item}")
                if st.form_submit_button("저장"):
                    if sel_site not in cl["site_enabled"]:
                        cl["site_enabled"][sel_site] = {}
                    for item, val in toggles.items():
                        cl["site_enabled"][sel_site][item] = val
                    save_data(data)
                    st.success("저장됨!")
                    st.rerun()

# ─────────────────────────────────────────
# ⚙️ 설정
# ─────────────────────────────────────────
elif menu == "⚙️ 설정":
    st.title("⚙️ 설정")
    st.markdown("---")
    tab_s1, tab_s2, tab_s3 = st.tabs(["🎨 화면 설정", "✏️ 이름 변경", "➕ 추가/삭제"])

    with tab_s1:
        with st.form("display_form"):
            new_sys = st.text_input("시스템 제목", value=settings["system_title"])
            new_side = st.text_input("사이드바 제목", value=settings["sidebar_title"])
            new_font = st.selectbox("글씨체", list(FONTS.keys()), index=list(FONTS.keys()).index(settings.get("font", "기본")))
            c1, c2 = st.columns(2)
            with c1:
                new_bg = st.color_picker("배경색", value=settings.get("bg_color", "#ffffff"))
            with c2:
                new_accent = st.color_picker("강조색", value=settings.get("accent_color", "#1f77b4"))
            if st.form_submit_button("저장", type="primary"):
                data["settings"].update({"system_title": new_sys, "sidebar_title": new_side, "font": new_font, "bg_color": new_bg, "accent_color": new_accent})
                save_data(data)
                st.success("✅ 저장됨! 새로고침하세요.")

    with tab_s2:
        st.markdown("### 고객사 이름 변경")
        for client in list(data["clients"].keys()):
            with st.form(f"rc_{client}"):
                new_name = st.text_input(f"👥 {client}", value=client)
                if st.form_submit_button("변경"):
                    if new_name and new_name != client:
                        data["clients"][new_name] = data["clients"].pop(client)
                        save_data(data)
                        st.success("변경됨! 새로고침하세요.")
        st.markdown("### 사업장 이름 변경")
        for client, cdata in data["clients"].items():
            for site in list(cdata["sites"].keys()):
                with st.form(f"rs_{client}_{site}"):
                    new_name = st.text_input(f"🍽️ {client} > {site}", value=site)
                    if st.form_submit_button("변경"):
                        if new_name and new_name != site:
                            cdata["sites"][new_name] = cdata["sites"].pop(site)
                            save_data(data)
                            st.success("변경됨! 새로고침하세요.")

    with tab_s3:
        st.markdown("### ➕ 고객사 추가")
        with st.form("ac"):
            nc = st.text_input("새 고객사 이름")
            if st.form_submit_button("추가"):
                if nc and nc not in data["clients"]:
                    data["clients"][nc] = {"sites": {}}
                    save_data(data)
                    st.success(f"'{nc}' 추가됨! 새로고침하세요.")
        st.markdown("### ➕ 사업장 추가")
        with st.form("as"):
            tc = st.selectbox("고객사", list(data["clients"].keys()))
            ns = st.text_input("새 사업장 이름")
            if st.form_submit_button("추가"):
                if ns and ns not in data["clients"][tc]["sites"]:
                    data["clients"][tc]["sites"][ns] = {"meetings": [], "issues": [], "history": []}
                    save_data(data)
                    st.success(f"'{ns}' 추가됨! 새로고침하세요.")
        st.markdown("### 🗑️ 삭제")
        with st.form("del"):
            dc = st.selectbox("삭제할 고객사", ["선택안함"] + list(data["clients"].keys()))
            all_sites = [(c, s) for c, cd in data["clients"].items() for s in cd["sites"].keys()]
            ds = st.selectbox("삭제할 사업장", ["선택안함"] + [f"{c} - {s}" for c, s in all_sites])
            if st.form_submit_button("삭제", type="primary"):
                if dc != "선택안함":
                    del data["clients"][dc]
                    save_data(data)
                    st.success("삭제됨! 새로고침하세요.")
                elif ds != "선택안함":
                    c, s = ds.split(" - ", 1)
                    del data["clients"][c]["sites"][s]
                    save_data(data)
                    st.success("삭제됨! 새로고침하세요.")
