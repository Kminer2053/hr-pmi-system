# -*- coding: utf-8 -*-
"""HR PMI Integrated Support System v4.0 - Streamlit Web App"""

import json
import base64
import calendar as cal_mod
import pandas as pd
import streamlit as st
from datetime import datetime, date, timedelta
from utils import (
    WORKFORCE_HEADERS, EMPLOYEE_HEADERS,
    read_any_file, read_xlsx_rows, read_csv_bytes,
    compare_regulations, find_duplicate_regulations,
    workforce_stats, workforce_group, to_number,
    normalize_workforce_rows, normalize_employee_rows,
    task_rate, generate_contract_docx, create_sample_contract_template,
    generate_report_docx, rows_to_csv_bytes,
)

APP_TITLE = "공공기관 HR PMI 통합지원 시스템 v4.0"


def init_state():
    defaults = {
        "project_name": "신규 통합 프로젝트",
        "inst_a_name": "A기관",
        "inst_b_name": "B기관",
        "merged_inst_name": "통합기관",
        "user_name": "팀원",
        "employees": [],
        "workforce": [],
        "contract_template_name": "",
        "contract_template_b64": "",
        "grade_maps": [],
        "welfare_maps": [],
        "risks": [
            ["고용승계 누락", "높음", "휴직자·계약직·파견인력 누락", "기준일 명부 3중 검증 및 예외자 법무검토", "인사/노무", "진행"],
            ["근로조건 불이익 변경", "높음", "보수·휴가·복지 하향 조정", "경과조치, 보전수당, 노사협의 및 동의절차 검토", "노무", "진행"],
            ["출신기관 갈등", "중간", "보직·승진·배치 불균형 인식", "배치기준 공개, 이의신청, 혼합형 TF 운영", "인사", "진행"],
        ],
        "tasks": [
            ["D-180", "통합추진단 구성 및 법률검토", "기획/인사", (date.today() + timedelta(days=7)).isoformat(), "진행"],
            ["D-120", "인력실사 및 고용승계 원칙 수립", "인사", (date.today() + timedelta(days=21)).isoformat(), "대기"],
            ["D-90", "직원 설명회 및 FAQ 운영", "인사/홍보", (date.today() + timedelta(days=35)).isoformat(), "대기"],
            ["D-30", "승계발령·급여·4대보험 이관 점검", "인사/재무", (date.today() + timedelta(days=60)).isoformat(), "대기"],
            ["D+100", "안정화 점검 및 이의신청 처리", "통합PMO", (date.today() + timedelta(days=160)).isoformat(), "대기"],
        ],
        "reg_rows": [],
        "duplicate_words": [],
        "duplicate_rules": [],
        "comments": [],
        "dashboard_fields": [
            {"title": "기관장 보고용 핵심메시지", "content": ""},
            {"title": "팀장 특이사항", "content": ""},
            {"title": "오늘의 점검사항", "content": ""},
            {"title": "추가 메모", "content": ""},
        ],
        "text_a": "",
        "text_b": "",
        "report_text": "",
        "cal_year": date.today().year,
        "cal_month": date.today().month,
        "workforce_base_year": date.today().year,
        "free_files": [],
        "team_projects": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ─── Helpers ───
def kpi_card(label, value, color="#1976D2"):
    st.markdown(
        f"""<div style="background:{color}11;border:1px solid {color}33;border-radius:10px;
        padding:12px 16px;text-align:center;">
        <div style="font-size:0.8rem;color:#666;">{label}</div>
        <div style="font-size:1.6rem;font-weight:700;color:{color};">{value}</div>
        </div>""", unsafe_allow_html=True)


def editable_table(data, columns, key, height=400):
    if not data:
        st.info("데이터가 없습니다.")
        return data
    df = pd.DataFrame(data, columns=columns)
    edited = st.data_editor(df, key=key, num_rows="dynamic", use_container_width=True, height=height)
    return edited.values.tolist()


def download_csv(data, columns, filename):
    rows = [columns] + data
    csv_bytes = rows_to_csv_bytes(rows)
    st.download_button(f"CSV 저장 ({filename})", csv_bytes, filename, "text/csv")


# ─── Tab 1: 프로젝트 ───
def tab_project():
    st.header("프로젝트 설정")
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.project_name = st.text_input("프로젝트명", st.session_state.project_name, key="inp_proj")
        st.session_state.inst_a_name = st.text_input("A기관명", st.session_state.inst_a_name, key="inp_a")
        st.session_state.inst_b_name = st.text_input("B기관명", st.session_state.inst_b_name, key="inp_b")
    with col2:
        st.session_state.merged_inst_name = st.text_input("통합기관명", st.session_state.merged_inst_name, key="inp_m")
        st.session_state.user_name = st.text_input("작성자", st.session_state.user_name, key="inp_user")

    st.divider()
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("새 프로젝트 초기화", type="secondary"):
            for k in ["employees", "workforce", "grade_maps", "welfare_maps", "risks", "tasks",
                       "reg_rows", "duplicate_words", "duplicate_rules", "comments", "free_files"]:
                st.session_state[k] = []
            st.session_state.contract_template_name = ""
            st.session_state.contract_template_b64 = ""
            st.session_state.text_a = ""
            st.session_state.text_b = ""
            st.session_state.report_text = ""
            st.rerun()
    with c2:
        data = project_data()
        st.download_button("프로젝트 JSON 저장", json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"),
                           f"{st.session_state.project_name}_HRPMI.json", "application/json")
    with c3:
        uploaded = st.file_uploader("프로젝트 JSON 불러오기", type=["json"], key="proj_upload")
        if uploaded:
            data = json.loads(uploaded.read().decode("utf-8"))
            apply_project_data(data)
            st.success("프로젝트를 불러왔습니다.")
            st.rerun()

    st.divider()
    st.markdown("""### 사용 방법
1. 프로젝트명, A기관명, B기관명, 통합기관명을 입력합니다.
2. 각 탭에서 자료를 입력하거나 CSV 업로드를 합니다.
3. 고용승계 탭에서 DOCX 계약서 양식을 올리고 직원별 계약서를 자동 생성합니다.
4. 인력구조 탭에서 직급·연령·근속·입사연도 분포를 분석합니다.
5. 추진일정은 목록과 월간 달력으로 함께 확인합니다.
6. 프로젝트 JSON 저장/불러오기로 데이터를 관리합니다.

> 모든 데이터는 브라우저 세션에 저장되며, JSON 파일로 내보내기/가져오기할 수 있습니다.
""")


# ─── Tab 2: 대시보드 ───
def tab_dashboard():
    st.header("대시보드")
    high = sum(1 for r in st.session_state.risks if len(r) > 1 and r[1] == "높음")
    diff = sum(1 for r in st.session_state.reg_rows if r and r[0] != "동일")
    rate = task_rate(st.session_state.tasks)

    cols = st.columns(6)
    labels = ["프로젝트", "고용승계", "인력구조", "고위험 리스크", "추진 진행률", "사규 차이"]
    values = [st.session_state.project_name, f"{len(st.session_state.employees)}명",
              f"{len(st.session_state.workforce)}명", f"{high}건", f"{rate}%", f"{diff}건"]
    colors = ["#1976D2", "#2E7D32", "#1565C0", "#C62828", "#F57C00", "#7B1FA2"]
    for col, label, val, color in zip(cols, labels, values, colors):
        with col:
            kpi_card(label, val, color)

    st.divider()
    st.subheader("자동 요약")
    lines = [
        f"**프로젝트:** {st.session_state.project_name}",
        f"- {st.session_state.inst_a_name} + {st.session_state.inst_b_name} → {st.session_state.merged_inst_name}",
        "", "**핵심 체크포인트**",
        "1. 고용승계: 휴직자, 계약직, 징계자, 산재자 누락 검증",
        "2. 인력구조: 직급·연령·근속·입사연도 분포와 조직형태 진단",
        "3. 사규비교: 차이 조항과 중복 규정·공통단어 확인",
        "4. 직급/복지 매핑: 원칙과 리스크를 함께 관리",
        "5. 추진일정: 목록과 월간 달력으로 상태 관리",
        "", "**주요 리스크**",
    ]
    for i, r in enumerate(st.session_state.risks[:8], 1):
        lines.append(f"{i}. **[{r[1] if len(r) > 1 else ''}]** {r[0] if r else ''} - {r[3] if len(r) > 3 else ''}")
    st.markdown("\n".join(lines))

    st.divider()
    st.subheader("사용자 입력 대시보드")
    fields = st.session_state.dashboard_fields
    for i, field in enumerate(fields):
        with st.expander(f"{field.get('title', '입력란')}", expanded=True):
            fields[i]["title"] = st.text_input("제목", field.get("title", ""), key=f"dash_title_{i}")
            fields[i]["content"] = st.text_area("내용", field.get("content", ""), key=f"dash_content_{i}", height=100)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("입력란 추가"):
            st.session_state.dashboard_fields.append({"title": "새 입력란", "content": ""})
            st.rerun()
    with c2:
        if len(fields) > 0 and st.button("마지막 입력란 삭제"):
            st.session_state.dashboard_fields.pop()
            st.rerun()


# ─── Tab 3: 고용승계·계약서 ───
def tab_employees():
    st.header("고용승계·계약서")

    with st.expander("개인별 계약서 자동작성", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            tpl = st.file_uploader("DOCX 계약서 양식 업로드", type=["docx"], key="contract_tpl")
            if tpl:
                st.session_state.contract_template_b64 = base64.b64encode(tpl.read()).decode("ascii")
                st.session_state.contract_template_name = tpl.name
                st.success(f"양식 등록: {tpl.name}")
        with c2:
            if st.session_state.contract_template_name:
                st.info(f"등록 양식: {st.session_state.contract_template_name}")
            else:
                st.warning("등록된 양식 없음")
        with c3:
            sample = create_sample_contract_template()
            st.download_button("계약서 양식 예시 다운로드", sample, "고용승계_계약서_양식_예시.docx",
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        st.caption("치환표시: {{성명}}, {{생년월일}}, {{소속}}, {{직급}}, {{월급여}}, {{통합기관명}} 등")

        if st.session_state.employees and st.session_state.contract_template_b64:
            emp_names = [f"{i}: {e[1] if len(e) > 1 else '직원'}" for i, e in enumerate(st.session_state.employees)]
            sel = st.selectbox("계약서 생성 대상", emp_names, key="contract_sel")
            if st.button("선택 직원 계약서 생성"):
                idx = int(sel.split(":")[0])
                emp = st.session_state.employees[idx]
                padded = (list(emp) + [""] * len(EMPLOYEE_HEADERS))[:len(EMPLOYEE_HEADERS)]
                mapping = dict(zip(EMPLOYEE_HEADERS, padded))
                mapping.update({
                    "프로젝트명": st.session_state.project_name,
                    "A기관명": st.session_state.inst_a_name,
                    "B기관명": st.session_state.inst_b_name,
                    "통합기관명": st.session_state.merged_inst_name,
                    "작성일": date.today().isoformat(),
                })
                tpl_bytes = base64.b64decode(st.session_state.contract_template_b64)
                result, hits = generate_contract_docx(tpl_bytes, mapping)
                name = mapping.get("성명", "직원")
                st.download_button("계약서 다운로드", result, f"{name}_고용승계_계약서.docx",
                                   "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    st.divider()
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("행 추가", key="emp_add"):
            st.session_state.employees.append(["", "", "", "", "", "", "정규직", "", "0", "", "", "승계", ""])
            st.rerun()
    with c2:
        emp_csv = st.file_uploader("CSV 업로드", type=["csv"], key="emp_csv")
        if emp_csv:
            rows = read_csv_bytes(emp_csv.read())
            if rows and any(h in rows[0] for h in EMPLOYEE_HEADERS):
                rows = rows[1:]
            st.session_state.employees = normalize_employee_rows(rows)
            st.success(f"{len(rows)}건 업로드 완료")
            st.rerun()
    with c3:
        if st.session_state.employees:
            download_csv(st.session_state.employees, EMPLOYEE_HEADERS, "고용승계_자료.csv")
    with c4:
        tpl_csv = rows_to_csv_bytes([EMPLOYEE_HEADERS, ["E001", "홍길동", "1990-01-01", "서울시", "인사팀", "4급", "정규직", "2020-01-01", "3500000", "2026-09-01", "", "승계", ""]])
        st.download_button("CSV 양식 다운로드", tpl_csv, "고용승계_CSV_양식.csv", "text/csv")

    st.session_state.employees = editable_table(
        st.session_state.employees, EMPLOYEE_HEADERS, "emp_editor")


# ─── Tab 4: 인력구조 분석 ───
def tab_workforce():
    st.header("인력구조 분석")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        wf = st.file_uploader("CSV/XLSX 업로드", type=["csv", "xlsx"], key="wf_upload")
        if wf:
            try:
                if wf.name.endswith(".xlsx"):
                    rows = read_xlsx_rows(wf.read())
                else:
                    rows = read_csv_bytes(wf.read())
                st.session_state.workforce = normalize_workforce_rows(rows)
                st.success(f"{len(st.session_state.workforce)}건 분석 완료")
                st.rerun()
            except Exception as e:
                st.error(str(e))
    with c2:
        tpl = rows_to_csv_bytes([WORKFORCE_HEADERS, ["E001", "1985", "", "4급", "4", "2010", "16", "본사", "관리", "Y"]])
        st.download_button("CSV 양식 저장", tpl, "인력구조_분석_양식.csv", "text/csv")
    with c3:
        if st.button("샘플 데이터 로드"):
            st.session_state.workforce = [
                ["E001", "1973", "", "2급", "2", "1998", "28", "본사", "관리", "Y"],
                ["E002", "1977", "", "3급", "3", "2002", "24", "서울본부", "영업", "Y"],
                ["E003", "1980", "", "3급", "3", "2005", "21", "부산본부", "영업", "Y"],
                ["E004", "1984", "", "4급", "4", "2009", "17", "서울본부", "영업", "Y"],
                ["E005", "1987", "", "4급", "4", "2012", "14", "대전본부", "물류", "N"],
                ["E006", "1991", "", "5급", "5", "2016", "10", "광주본부", "영업", "N"],
                ["E007", "1995", "", "6급", "6", "2020", "6", "부산본부", "영업", "N"],
                ["E008", "1998", "", "6급", "6", "2023", "3", "대전본부", "지원", "N"],
            ]
            st.rerun()
    with c4:
        st.session_state.workforce_base_year = st.number_input("나이 기준연도", 2000, 2100,
                                                                st.session_state.workforce_base_year, key="wf_year")

    if not st.session_state.workforce:
        st.info("CSV 또는 XLSX 인력정보를 업로드하세요.")
        return

    stats = workforce_stats(st.session_state.workforce, st.session_state.workforce_base_year)
    if stats:
        st.session_state.workforce = stats["normalized"]
        cols = st.columns(6)
        kpi_labels = ["총 인원", "평균 나이", "평균 근속", "중간직급 비중", "관리자 비중", "구조 판정"]
        kpi_values = [f"{stats['n']}명", f"{stats['avg_age']:.1f}세", f"{stats['avg_tenure']:.1f}년",
                      f"{stats['mid_share'] * 100:.1f}%", f"{stats['manager_share'] * 100:.1f}%", stats["shape"]]
        kpi_colors = ["#1976D2", "#2E7D32", "#1565C0", "#F57C00", "#7B1FA2", "#C62828"]
        for col, lbl, val, clr in zip(cols, kpi_labels, kpi_values, kpi_colors):
            with col:
                kpi_card(lbl, val, clr)

        st.divider()
        c_left, c_right = st.columns(2)
        with c_left:
            st.subheader("조직구조 진단")
            advice = ("중간직급 기반이 확보되어 있습니다." if stats["shape"] == "항아리형"
                      else "하위직급 대비 중간직급이 얇습니다." if stats["shape"] == "피라미드형"
                      else "직급별 역할·관리폭과 자연감소 전망을 함께 점검하세요.")
            st.markdown(f"""
- **구조유형:** {stats['shape']}
- 상위 {stats['bands']['상위']}명 / 중간 {stats['bands']['중간']}명 / 하위 {stats['bands']['하위']}명
- 중간직급 비중 {stats['mid_share'] * 100:.1f}%, 관리자 비중 {stats['manager_share'] * 100:.1f}%

**검토의견:** {advice}
""")

        with c_right:
            st.subheader("분포 분석")
            wf = st.session_state.workforce
            dist_tab = st.tabs(["직급", "연령", "근속", "입사연도"])
            with dist_tab[0]:
                g = workforce_group(wf, 3)
                st.dataframe(pd.DataFrame([(k, v, f"{v / stats['n'] * 100:.1f}%") for k, v in g.items()],
                                          columns=["구간", "인원", "비중"]), use_container_width=True, hide_index=True)
            with dist_tab[1]:
                g = workforce_group(wf, 2, lambda r: "20대" if to_number(r[2]) < 30 else "30대" if to_number(r[2]) < 40 else "40대" if to_number(r[2]) < 50 else "50대" if to_number(r[2]) < 60 else "60대+")
                st.dataframe(pd.DataFrame([(k, v, f"{v / stats['n'] * 100:.1f}%") for k, v in g.items()],
                                          columns=["구간", "인원", "비중"]), use_container_width=True, hide_index=True)
            with dist_tab[2]:
                g = workforce_group(wf, 6, lambda r: "5년 미만" if to_number(r[6]) < 5 else "5~9년" if to_number(r[6]) < 10 else "10~14년" if to_number(r[6]) < 15 else "15~19년" if to_number(r[6]) < 20 else "20년+")
                st.dataframe(pd.DataFrame([(k, v, f"{v / stats['n'] * 100:.1f}%") for k, v in g.items()],
                                          columns=["구간", "인원", "비중"]), use_container_width=True, hide_index=True)
            with dist_tab[3]:
                g = workforce_group(wf, 5, lambda r: "1999년 이전" if to_number(r[5]) < 2000 else "2000년대" if to_number(r[5]) < 2010 else "2010년대" if to_number(r[5]) < 2020 else "2020년대")
                st.dataframe(pd.DataFrame([(k, v, f"{v / stats['n'] * 100:.1f}%") for k, v in g.items()],
                                          columns=["구간", "인원", "비중"]), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("인력 현황 상세")
    if st.session_state.workforce:
        df = pd.DataFrame(st.session_state.workforce, columns=WORKFORCE_HEADERS)
        st.dataframe(df, use_container_width=True, hide_index=True, height=400)
        download_csv(st.session_state.workforce, WORKFORCE_HEADERS, "인력구조_분석결과.csv")


# ─── Tab 5: 사규 비교 ───
def tab_regulation():
    st.header("사규 비교")
    a_name = st.session_state.inst_a_name
    b_name = st.session_state.inst_b_name

    c1, c2 = st.columns(2)
    with c1:
        fa = st.file_uploader(f"{a_name} 사규 업로드", type=["txt", "pdf", "docx", "hwpx"], key="reg_a")
        if fa:
            st.session_state.text_a = read_any_file(fa.name, fa.read())
            st.success(f"{a_name} 사규 로드 완료")
    with c2:
        fb = st.file_uploader(f"{b_name} 사규 업로드", type=["txt", "pdf", "docx", "hwpx"], key="reg_b")
        if fb:
            st.session_state.text_b = read_any_file(fb.name, fb.read())
            st.success(f"{b_name} 사규 로드 완료")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.session_state.text_a = st.text_area(f"{a_name} 사규", st.session_state.text_a, height=250, key="ta_a")
    with col2:
        st.session_state.text_b = st.text_area(f"{b_name} 사규", st.session_state.text_b, height=250, key="ta_b")
    with col3:
        st.markdown("**차이점 비교 결과**")
        diff_placeholder = st.empty()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        run_compare = st.button("비교 실행", type="primary")
    with c2:
        run_dup = st.button("중복 규정·단어 찾기")
    with c3:
        if st.session_state.reg_rows:
            changed = [r for r in st.session_state.reg_rows if r[0] != "동일"]
            download_csv(changed, ["구분", "조문", "요약", "리스크", f"{a_name} 원문", f"{b_name} 원문"], "사규_비교표.csv")
    with c4:
        if st.session_state.duplicate_words or st.session_state.duplicate_rules:
            dup_rows = [["구분", "항목1", "항목2", "A빈도/유사도", "B빈도", "합계/내용"]]
            for w, ca, cb, total in st.session_state.duplicate_words:
                dup_rows.append(["공통단어", w, "", ca, cb, total])
            for ak, bk, score, at, bt in st.session_state.duplicate_rules:
                dup_rows.append(["유사규정", ak, bk, score, at, bt])
            csv_b = rows_to_csv_bytes(dup_rows)
            st.download_button("중복검색 CSV 저장", csv_b, "사규_중복규정_공통단어.csv", "text/csv")

    if run_compare:
        if not st.session_state.text_a.strip() or not st.session_state.text_b.strip():
            st.warning("두 기관 사규를 모두 입력하세요.")
        else:
            st.session_state.reg_rows = compare_regulations(st.session_state.text_a, st.session_state.text_b)
            st.success(f"비교 완료: 전체 {len(st.session_state.reg_rows)}건")

    if run_dup:
        if not st.session_state.text_a.strip() or not st.session_state.text_b.strip():
            st.warning("두 기관 사규를 모두 입력하세요.")
        else:
            with st.spinner("중복 검색 중..."):
                st.session_state.duplicate_words, st.session_state.duplicate_rules = find_duplicate_regulations(
                    st.session_state.text_a, st.session_state.text_b)
            st.success(f"공통 단어 {len(st.session_state.duplicate_words)}개, 유사 규정 {len(st.session_state.duplicate_rules)}쌍")

    if st.session_state.reg_rows:
        changed = [r for r in st.session_state.reg_rows if r[0] != "동일"]
        with diff_placeholder.container():
            st.markdown(f"전체 {len(st.session_state.reg_rows)}건 / **차이 {len(changed)}건**")
        if changed:
            st.divider()
            st.subheader("차이 조항 상세")
            for i, r in enumerate(changed, 1):
                with st.expander(f"{i}. [{r[0]}] {r[1]} — {r[3]}"):
                    st.markdown(f"**요약:** {r[2]}")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown(f"**{a_name}**")
                        st.text(r[4][:1500] if r[4] else "(없음)")
                    with c2:
                        st.markdown(f"**{b_name}**")
                        st.text(r[5][:1500] if r[5] else "(없음)")

    if st.session_state.duplicate_words or st.session_state.duplicate_rules:
        st.divider()
        st.subheader("중복 검색 결과")
        t1, t2 = st.tabs(["공통 핵심단어", "유사·중복 규정"])
        with t1:
            if st.session_state.duplicate_words:
                df = pd.DataFrame(st.session_state.duplicate_words[:100], columns=["단어", a_name, b_name, "합계"])
                st.dataframe(df, use_container_width=True, hide_index=True)
        with t2:
            for i, row in enumerate(st.session_state.duplicate_rules[:50], 1):
                with st.expander(f"{i}. {a_name} {row[0]} ↔ {b_name} {row[1]} / 유사도 {row[2]}%"):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.text(row[3])
                    with c2:
                        st.text(row[4])


# ─── Tab 6: 직급 매핑 ───
def tab_grade_mapping():
    st.header("직급 매핑")
    a = st.session_state.inst_a_name
    b = st.session_state.inst_b_name
    m = st.session_state.merged_inst_name
    cols = [a, b, f"{m} 통합직급", "원칙", "리스크"]

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("행 추가", key="grade_add"):
            st.session_state.grade_maps.append([""] * 5)
            st.rerun()
    with c2:
        tpl = rows_to_csv_bytes([cols, [f"{a} 4급", f"{b} 4급", f"{m} 4급", "동일직급 매핑", "승진연한 차이 검토"]])
        st.download_button("CSV 양식", tpl, "직급_CSV_양식.csv", "text/csv")
    with c3:
        gc = st.file_uploader("CSV 업로드", type=["csv"], key="grade_csv")
        if gc:
            rows = read_csv_bytes(gc.read())
            if rows and any(h in rows[0] for h in cols):
                rows = rows[1:]
            st.session_state.grade_maps = [(r + [""] * 5)[:5] for r in rows]
            st.rerun()

    st.session_state.grade_maps = editable_table(st.session_state.grade_maps, cols, "grade_editor")
    if st.session_state.grade_maps:
        download_csv(st.session_state.grade_maps, cols, "직급_매핑_자료.csv")


# ─── Tab 7: 복지 매핑 ───
def tab_welfare_mapping():
    st.header("복지 매핑")
    a = st.session_state.inst_a_name
    b = st.session_state.inst_b_name
    cols = ["항목", a, b, "원칙", "리스크"]

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("행 추가", key="welfare_add"):
            st.session_state.welfare_maps.append([""] * 5)
            st.rerun()
    with c2:
        tpl = rows_to_csv_bytes([cols, ["복지포인트", "100만원", "120만원", "재원 범위 내 단계적 조정", "예산·형평성 검토"]])
        st.download_button("CSV 양식", tpl, "복지_CSV_양식.csv", "text/csv")
    with c3:
        wc = st.file_uploader("CSV 업로드", type=["csv"], key="welfare_csv")
        if wc:
            rows = read_csv_bytes(wc.read())
            if rows and any(h in rows[0] for h in cols):
                rows = rows[1:]
            st.session_state.welfare_maps = [(r + [""] * 5)[:5] for r in rows]
            st.rerun()

    st.session_state.welfare_maps = editable_table(st.session_state.welfare_maps, cols, "welfare_editor")
    if st.session_state.welfare_maps:
        download_csv(st.session_state.welfare_maps, cols, "복지_매핑_자료.csv")


# ─── Tab 8: 리스크 ───
def tab_risks():
    st.header("리스크 관리")
    cols = ["리스크", "수준", "원인", "대응방안", "담당", "상태"]

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("행 추가", key="risk_add"):
            st.session_state.risks.append(["", "중간", "", "", "", "대기"])
            st.rerun()
    with c2:
        tpl = rows_to_csv_bytes([cols, ["근로조건 불이익 변경", "높음", "보수·휴가·복지 하향", "경과조치 및 노사협의", "노무", "진행"]])
        st.download_button("CSV 양식", tpl, "리스크_CSV_양식.csv", "text/csv")
    with c3:
        rc = st.file_uploader("CSV 업로드", type=["csv"], key="risk_csv")
        if rc:
            rows = read_csv_bytes(rc.read())
            if rows and any(h in rows[0] for h in cols):
                rows = rows[1:]
            st.session_state.risks = [(r + [""] * 6)[:6] for r in rows]
            st.rerun()

    st.session_state.risks = editable_table(st.session_state.risks, cols, "risk_editor")
    if st.session_state.risks:
        download_csv(st.session_state.risks, cols, "리스크_자료.csv")


# ─── Tab 9: 추진일정 ───
def tab_schedule():
    st.header("추진일정")
    cols_def = ["시기", "과제", "담당", "기한", "상태"]
    rate = task_rate(st.session_state.tasks)
    st.progress(rate / 100, text=f"진행률: {rate}%")

    tab_list, tab_cal = st.tabs(["목록 보기", "달력 보기"])

    with tab_list:
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("행 추가", key="task_add"):
                st.session_state.tasks.append(["", "", "", "", "대기"])
                st.rerun()
        with c2:
            tpl = rows_to_csv_bytes([cols_def, ["D-120", "인력실사 및 고용승계 원칙 수립", "인사", "2026-08-31", "진행"]])
            st.download_button("CSV 양식", tpl, "추진일정_CSV_양식.csv", "text/csv")
        with c3:
            tc = st.file_uploader("CSV 업로드", type=["csv"], key="task_csv")
            if tc:
                rows = read_csv_bytes(tc.read())
                if rows and any(h in rows[0] for h in cols_def):
                    rows = rows[1:]
                st.session_state.tasks = [(r + [""] * 5)[:5] for r in rows]
                st.rerun()

        st.session_state.tasks = editable_table(st.session_state.tasks, cols_def, "task_editor")
        if st.session_state.tasks:
            download_csv(st.session_state.tasks, cols_def, "추진일정_자료.csv")

    with tab_cal:
        c1, c2, c3, c4 = st.columns([1, 1, 1, 3])
        with c1:
            if st.button("◀ 이전달"):
                v = st.session_state.cal_year * 12 + st.session_state.cal_month - 2
                st.session_state.cal_year, mi = divmod(v, 12)
                st.session_state.cal_month = mi + 1
                st.rerun()
        with c2:
            if st.button("오늘"):
                st.session_state.cal_year = date.today().year
                st.session_state.cal_month = date.today().month
                st.rerun()
        with c3:
            if st.button("다음달 ▶"):
                v = st.session_state.cal_year * 12 + st.session_state.cal_month
                st.session_state.cal_year, mi = divmod(v, 12)
                st.session_state.cal_month = mi + 1
                st.rerun()
        with c4:
            st.markdown(f"### {st.session_state.cal_year}년 {st.session_state.cal_month}월")

        render_calendar()


def render_calendar():
    y, m = st.session_state.cal_year, st.session_state.cal_month
    weeks = cal_mod.Calendar(firstweekday=0).monthdayscalendar(y, m)
    while len(weeks) < 6:
        weeks.append([0] * 7)

    by_day = {}
    for task in st.session_state.tasks:
        raw = str(task[3] if len(task) > 3 else "").strip().replace(".", "-").replace("/", "-")
        for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M"):
            try:
                due = datetime.strptime(raw, fmt).date()
                if due.year == y and due.month == m:
                    by_day.setdefault(due.day, []).append(task)
                break
            except Exception:
                pass

    today = date.today()
    day_names = st.columns(7)
    for i, name in enumerate(["월", "화", "수", "목", "금", "토", "일"]):
        color = "#c62828" if i == 6 else "#235a9f" if i == 5 else "#333"
        day_names[i].markdown(f"<div style='text-align:center;font-weight:700;color:{color};'>{name}</div>",
                              unsafe_allow_html=True)

    for week in weeks:
        cols = st.columns(7)
        for i, day in enumerate(week):
            with cols[i]:
                if day == 0:
                    st.markdown("<div style='height:80px;'></div>", unsafe_allow_html=True)
                else:
                    tasks = by_day.get(day, [])
                    is_today = (y, m, day) == (today.year, today.month, today.day)
                    bg = "#fff4c7" if is_today else "#f8f9fa"
                    entries = []
                    for t in tasks[:3]:
                        status = t[4] if len(t) > 4 else ""
                        mark = "&#10003;" if status == "완료" else "&#9679;" if status == "진행" else "&#9675;"
                        title = str(t[1] if len(t) > 1 else "")[:10]
                        entries.append(f"<div style='font-size:0.7rem;'>{mark} {title}</div>")
                    if len(tasks) > 3:
                        entries.append(f"<div style='font-size:0.65rem;color:#888;'>외 {len(tasks) - 3}건</div>")
                    html = f"""<div style='background:{bg};border:1px solid #ddd;border-radius:6px;
                    padding:4px;min-height:80px;'>
                    <div style='font-weight:600;font-size:0.85rem;'>{day}</div>
                    {''.join(entries)}</div>"""
                    st.markdown(html, unsafe_allow_html=True)


# ─── Tab 10: 보고서 ───
def tab_report():
    st.header("보고서")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("기관장 보고서"):
            st.session_state.report_text = report_exec()
    with c2:
        if st.button("노조 설명자료"):
            st.session_state.report_text = report_union()
    with c3:
        if st.button("주무부처 보고자료"):
            st.session_state.report_text = report_ministry()
    with c4:
        csv_file = st.file_uploader("CSV→보고서 변환", type=["csv"], key="report_csv")
        if csv_file:
            st.session_state.report_text = csv_to_report(csv_file)

    st.session_state.report_text = st.text_area("보고서 내용", st.session_state.report_text, height=500, key="report_area")

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.session_state.report_text:
            st.download_button("TXT 저장", st.session_state.report_text.encode("utf-8"),
                               "HR_PMI_보고서.txt", "text/plain")
    with c2:
        if st.session_state.report_text:
            docx_bytes = generate_report_docx(st.session_state.report_text)
            st.download_button("DOCX 저장", docx_bytes, "HR_PMI_보고서.docx",
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with c3:
        if st.button("인력구조 보고문 생성"):
            stats = workforce_stats(st.session_state.workforce, st.session_state.workforce_base_year)
            if stats:
                st.session_state.report_text = report_workforce(stats)
            else:
                st.warning("먼저 인력정보를 업로드하세요.")


def report_exec():
    s = st.session_state
    high = sum(1 for r in s.risks if len(r) > 1 and r[1] == "높음")
    diff = sum(1 for r in s.reg_rows if r[0] != "동일")
    dash_lines = "\n".join(
        f"- {f.get('title', '')}: {f.get('content', '') or '내용 없음'}"
        for f in s.dashboard_fields if f.get("title") or f.get("content")
    ) or "- 별도 입력사항 없음"
    return f"""기관장 보고자료

제목: {s.project_name} HR PMI 추진현황

1. 추진배경
{s.inst_a_name}와 {s.inst_b_name}의 통합에 따라 고용승계, 사규비교, 직급·복지 매핑, 조직 안정화, 노사협의 등 인사 분야 PMI 과제의 체계적 관리가 필요합니다.

2. 현재 현황
- 고용승계 관리대상: {len(s.employees)}명
- 인력구조 분석대상: {len(s.workforce)}명
- 사규 차이 검토조문: {diff}건
- 고위험 리스크: {high}건
- 추진과제 진행률: {task_rate(s.tasks)}%

3. 대시보드 입력사항
{dash_lines}

4. 향후 조치
고용승계 대상자 확정, 불이익 변경 가능 조항 노무·법무 검토, 직급·복지 경과조치 마련, 통합 후 100일 안정화 점검을 추진하겠습니다.
"""


def report_union():
    s = st.session_state
    return f"""노동조합 설명자료

1. 고용승계 원칙
{s.inst_a_name} 및 {s.inst_b_name}의 통합 과정에서 통합일 현재 재직 중인 직원은 원칙적으로 고용승계 대상으로 관리합니다.

2. 근로조건 처리방향
임금, 복무, 휴가, 복리후생 등 근로조건은 급격한 저하가 발생하지 않도록 경과조치를 우선 검토합니다.

3. 협의 필요사항
불이익 변경 가능 조항, 직급·보수 매핑, 복지제도 조정, 전보·배치 기준은 노사협의 절차에 따라 검토합니다.
"""


def report_ministry():
    s = st.session_state
    return f"""주무부처 보고자료

1. 보고요지
{s.project_name} 추진에 따라 {s.inst_a_name}와 {s.inst_b_name}의 인사·노무 PMI 과제를 체계적으로 관리 중입니다.

2. 주요 관리현황
- 승계대상 관리인원: {len(s.employees)}명
- 인력구조 분석인원: {len(s.workforce)}명
- 고위험 리스크: {sum(1 for r in s.risks if len(r) > 1 and r[1] == '높음')}건
- 사규 차이 검토건수: {sum(1 for r in s.reg_rows if r[0] != '동일')}건
- 추진과제 진행률: {task_rate(s.tasks)}%

3. 협조 필요사항
정원·인건비·조직개편 관련 협의, 제도정비 관련 승인절차, 통합 후 안정화 기간에 대한 정책적 지원이 필요합니다.
"""


def report_workforce(stats):
    return f"""인력구조 분석 보고

1. 분석 개요
- 대상인원: {stats['n']}명
- 평균연령: {stats['avg_age']:.1f}세 / 평균근속: {stats['avg_tenure']:.1f}년

2. 직급구조 진단
- 구조유형: {stats['shape']}
- 상위직급 {stats['bands']['상위']}명, 중간직급 {stats['bands']['중간']}명, 하위직급 {stats['bands']['하위']}명
- 중간직급 비중: {stats['mid_share'] * 100:.1f}% / 관리자 비중: {stats['manager_share'] * 100:.1f}%

3. 검토방향
인력구조 분석 결과를 관리폭, 사업장·거점 수, 매출책임, 업무난도 및 향후 자연감소 전망과 결합하여 통합기관의 적정 직급·정원 구조를 설계할 필요가 있습니다.
"""


def csv_to_report(csv_file):
    rows = read_csv_bytes(csv_file.read())
    if not rows:
        return ""
    headers, body = rows[0], rows[1:]
    lines = [f"{csv_file.name.rsplit('.', 1)[0]} 검토 보고서", "", "1. 검토 개요",
             f"- 자료건수: {len(body)}건", f"- 주요컬럼: {', '.join(headers)}", "", "2. 세부 검토내용"]
    for i, row in enumerate(body[:50], 1):
        lines.append(f"{i}) " + " / ".join(
            [f"{headers[j] if j < len(headers) else '항목' + str(j + 1)}: {row[j]}" for j in range(len(row))]))
    lines += ["", "3. 검토의견",
              "자료상 주요 항목을 기준으로 통합 전후 차이, 인력 영향, 노사협의 필요사항, 후속 검증 필요사항을 추가 확인할 필요가 있습니다."]
    return "\n".join(lines)


# ─── Tab 11: 자유자료 공유 ───
def tab_free_share():
    st.header("자유자료 공유")
    with st.expander("자료 업로드", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            title = st.text_input("자료명", key="free_title")
            category = st.text_input("분류", "참고자료", key="free_cat")
        with c2:
            memo = st.text_area("메모", key="free_memo", height=100)
        uploaded = st.file_uploader("파일 선택", key="free_file")
        if uploaded and st.button("업로드·등록"):
            rec = {
                "title": title or uploaded.name,
                "category": category,
                "author": st.session_state.user_name,
                "uploaded_at": datetime.now().isoformat(timespec="seconds"),
                "original_name": uploaded.name,
                "memo": memo,
                "data_b64": base64.b64encode(uploaded.read()).decode("ascii"),
            }
            st.session_state.free_files.append(rec)
            st.success("자료를 등록했습니다.")
            st.rerun()

    if st.session_state.free_files:
        st.divider()
        st.subheader("등록 자료 목록")
        for i, r in enumerate(reversed(st.session_state.free_files)):
            with st.expander(f"{r['title']} ({r['category']}) — {r['author']} / {r['uploaded_at']}"):
                st.markdown(f"**원본파일:** {r['original_name']}")
                if r.get("memo"):
                    st.markdown(f"**메모:** {r['memo']}")
                data = base64.b64decode(r["data_b64"])
                st.download_button("파일 다운로드", data, r["original_name"], key=f"free_dl_{i}")


# ─── Tab 12: 댓글/검토의견 ───
def tab_comments():
    st.header("댓글/검토의견")
    with st.expander("댓글 작성", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            target = st.selectbox("대상 탭", ["공통", "대시보드", "고용승계·계약서", "인력구조 분석", "사규 비교",
                                              "직급 매핑", "복지 매핑", "리스크", "추진일정", "보고서",
                                              "자유자료 공유", "팀장 상황판"], key="cmt_target")
        with c2:
            ctype = st.selectbox("유형", ["검토의견", "수정요청", "확인완료", "질문", "답변"], key="cmt_type")
        body = st.text_area("내용", key="cmt_body")
        if st.button("댓글 추가") and body.strip():
            st.session_state.comments.append([
                datetime.now().isoformat(timespec="seconds"),
                st.session_state.user_name, target, ctype, body.strip(), "등록"
            ])
            st.success("댓글을 추가했습니다.")
            st.rerun()

    if st.session_state.comments:
        st.divider()
        cols = ["일시", "작성자", "대상", "유형", "내용", "상태"]
        st.session_state.comments = editable_table(st.session_state.comments, cols, "cmt_editor")
        download_csv(st.session_state.comments, cols, "댓글_검토의견.csv")


# ─── Tab 13: 팀 공유 ───
def tab_team_share():
    st.header("팀 공유")
    st.markdown("""
> **웹 버전 팀 공유 방식:** 프로젝트 JSON 파일을 다운로드/업로드하여 팀원 간 데이터를 공유합니다.
""")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("내 프로젝트 공유")
        data = project_data()
        author = st.session_state.user_name or "팀원"
        proj = st.session_state.project_name or "프로젝트"
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button("프로젝트 JSON 다운로드 (공유용)", json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"),
                           f"HR_PMI_{proj}_{author}_{stamp}.json", "application/json")

    with c2:
        st.subheader("팀원 프로젝트 불러오기")
        files = st.file_uploader("팀원 프로젝트 JSON 업로드 (복수 선택 가능)", type=["json"],
                                 accept_multiple_files=True, key="team_upload")
        if files:
            for f in files:
                data = json.loads(f.read().decode("utf-8"))
                data["_file"] = f.name
                st.session_state.team_projects.append(data)
            st.success(f"{len(files)}개 프로젝트를 로드했습니다.")
            st.rerun()

    if st.session_state.team_projects:
        st.divider()
        st.subheader("로드된 팀원 프로젝트")
        for i, p in enumerate(st.session_state.team_projects):
            with st.expander(f"{p.get('project_name', '프로젝트')} — {p.get('author', '팀원')} ({p.get('saved_at', '')})"):
                st.markdown(f"- 고용승계: {len(p.get('employees', []))}건")
                st.markdown(f"- 인력구조: {len(p.get('workforce', []))}건")
                st.markdown(f"- 리스크: {len(p.get('risks', []))}건")
                st.markdown(f"- 추진일정: {len(p.get('tasks', []))}건")
                if st.button(f"이 프로젝트 불러오기", key=f"load_team_{i}"):
                    apply_project_data(p)
                    st.success("프로젝트를 불러왔습니다.")
                    st.rerun()


# ─── Tab 14: 팀장 상황판 ───
def tab_manager_dashboard():
    st.header("팀장 상황판")
    projects = st.session_state.team_projects

    if not projects:
        st.info("'팀 공유' 탭에서 팀원 프로젝트 JSON을 업로드하면 통합 상황판이 표시됩니다.")
        return

    by_author = {}
    all_tasks = []
    all_risks = []
    for p in projects:
        a = p.get("author", "팀원")
        info = by_author.setdefault(a, {"count": 0, "latest": "", "employees": 0, "workforce": 0,
                                         "grades": 0, "welfare": 0, "risks": 0, "tasks": 0, "rates": []})
        info["count"] += 1
        saved = p.get("saved_at", "")
        if saved > info["latest"]:
            info["latest"] = saved
        info["employees"] += len(p.get("employees", []))
        info["workforce"] += len(p.get("workforce", []))
        info["grades"] += len(p.get("grade_maps", []))
        info["welfare"] += len(p.get("welfare_maps", []))
        info["risks"] += len(p.get("risks", []))
        info["tasks"] += len(p.get("tasks", []))
        info["rates"].append(task_rate(p.get("tasks", [])))
        for t in p.get("tasks", []):
            all_tasks.append([a] + (t + [""] * 5)[:5])
        for r in p.get("risks", []):
            all_risks.append([a] + (r + [""] * 6)[:6])

    high = sum(1 for r in all_risks if len(r) > 2 and r[2] == "높음")
    rates = [task_rate(p.get("tasks", [])) for p in projects if p.get("tasks")]
    avg_rate = round(sum(rates) / len(rates)) if rates else 0

    cols = st.columns(5)
    kpi_l = ["공유 프로젝트", "작성 팀원", "평균 진행률", "고위험 리스크", "자유자료"]
    kpi_v = [len(projects), len(by_author), f"{avg_rate}%", high, len(st.session_state.free_files)]
    kpi_c = ["#1976D2", "#2E7D32", "#F57C00", "#C62828", "#7B1FA2"]
    for col, lbl, val, clr in zip(cols, kpi_l, kpi_v, kpi_c):
        with col:
            kpi_card(lbl, val, clr)

    st.divider()
    t1, t2, t3 = st.tabs(["팀원별 현황", "팀 전체 추진일정", "팀 전체 리스크"])
    with t1:
        rows = []
        for a, info in by_author.items():
            avg = round(sum(info["rates"]) / len(info["rates"])) if info["rates"] else 0
            rows.append([a, info["latest"], info["count"], info["employees"], info["workforce"],
                         info["grades"], info["welfare"], info["risks"], info["tasks"], f"{avg}%"])
        df = pd.DataFrame(rows, columns=["작성자", "최근저장", "프로젝트수", "고용승계", "인력구조",
                                          "직급", "복지", "리스크", "추진일정", "진행률"])
        st.dataframe(df, use_container_width=True, hide_index=True)
    with t2:
        if all_tasks:
            df = pd.DataFrame(all_tasks, columns=["작성자", "시기", "과제", "담당", "기한", "상태"])
            st.dataframe(df, use_container_width=True, hide_index=True)
    with t3:
        if all_risks:
            df = pd.DataFrame(all_risks, columns=["작성자", "리스크", "수준", "원인", "대응방안", "담당", "상태"])
            st.dataframe(df, use_container_width=True, hide_index=True)

    st.divider()
    if st.button("상황판 보고서 생성"):
        authors = sorted(by_author.keys())
        lines = ["팀장용 HR PMI 통합 상황판 보고", "",
                 "1. 총괄 현황",
                 f"- 공유 프로젝트 수: {len(projects)}건",
                 f"- 작성 팀원 수: {len(authors)}명",
                 f"- 고위험 리스크 수: {high}건",
                 f"- 자유자료 공유 건수: {len(st.session_state.free_files)}건", "",
                 "2. 팀원별 작성 현황"]
        for a in authors:
            ps = [p for p in projects if p.get("author", "팀원") == a]
            lines.append(f"- {a}: 프로젝트 {len(ps)}건, 고용승계 {sum(len(p.get('employees', [])) for p in ps)}건")
        lines += ["", "3. 고위험 리스크"]
        high_risks = [r for r in all_risks if len(r) > 2 and r[2] == "높음"]
        for i, r in enumerate(high_risks[:20], 1):
            row = (r + [""] * 7)[:7]
            lines.append(f"{i}) [{row[0]}] {row[1]} / 대응: {row[4]} / 상태: {row[6]}")
        st.session_state.report_text = "\n".join(lines)
        st.success("보고서 탭에서 확인하세요.")


# ─── Project data ───
def project_data():
    s = st.session_state
    return {
        "app_version": "4.0",
        "project_name": s.project_name,
        "inst_a_name": s.inst_a_name,
        "inst_b_name": s.inst_b_name,
        "merged_inst_name": s.merged_inst_name,
        "author": s.user_name,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "employees": s.employees,
        "workforce": s.workforce,
        "workforce_base_year": s.workforce_base_year,
        "contract_template_name": s.contract_template_name,
        "contract_template_b64": s.contract_template_b64,
        "grade_maps": s.grade_maps,
        "welfare_maps": s.welfare_maps,
        "risks": s.risks,
        "tasks": s.tasks,
        "reg_rows": s.reg_rows,
        "duplicate_words": s.duplicate_words,
        "duplicate_rules": s.duplicate_rules,
        "comments": s.comments,
        "dashboard_fields": s.dashboard_fields,
        "text_a": s.text_a,
        "text_b": s.text_b,
        "report": s.report_text,
        "free_files": [
            {k: v for k, v in f.items() if k != "data_b64"} for f in s.free_files
        ],
    }


def apply_project_data(data):
    s = st.session_state
    s.project_name = data.get("project_name", s.project_name)
    s.inst_a_name = data.get("inst_a_name", s.inst_a_name)
    s.inst_b_name = data.get("inst_b_name", s.inst_b_name)
    s.merged_inst_name = data.get("merged_inst_name", s.merged_inst_name)
    s.user_name = data.get("author", s.user_name)
    s.employees = normalize_employee_rows(data.get("employees", []))
    s.workforce = [(list(r) + [""] * len(WORKFORCE_HEADERS))[:len(WORKFORCE_HEADERS)] for r in data.get("workforce", [])]
    s.contract_template_name = data.get("contract_template_name", "")
    s.contract_template_b64 = data.get("contract_template_b64", "")
    s.grade_maps = data.get("grade_maps", [])
    s.welfare_maps = data.get("welfare_maps", [])
    s.risks = data.get("risks", [])
    s.tasks = data.get("tasks", [])
    s.reg_rows = data.get("reg_rows", [])
    s.duplicate_words = data.get("duplicate_words", [])
    s.duplicate_rules = data.get("duplicate_rules", [])
    s.comments = data.get("comments", [])
    s.dashboard_fields = data.get("dashboard_fields", s.dashboard_fields)
    s.text_a = data.get("text_a", "")
    s.text_b = data.get("text_b", "")
    s.report_text = data.get("report", "")
    try:
        s.workforce_base_year = int(data.get("workforce_base_year", date.today().year))
    except Exception:
        s.workforce_base_year = date.today().year


# ─── Main ───
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="📋", layout="wide", initial_sidebar_state="expanded")
    init_state()

    with st.sidebar:
        st.markdown(f"## HR PMI v4.0")
        st.caption(f"작성자: {st.session_state.user_name}")
        st.divider()
        menu = st.radio("메뉴", [
            "① 프로젝트",
            "② 대시보드",
            "③ 고용승계·계약서",
            "④ 인력구조 분석",
            "⑤ 사규 비교",
            "⑥ 직급 매핑",
            "⑦ 복지 매핑",
            "⑧ 리스크",
            "⑨ 추진일정",
            "⑩ 보고서",
            "⑪ 자유자료 공유",
            "⑫ 댓글/검토의견",
            "⑬ 팀 공유",
            "⑭ 팀장 상황판",
        ], label_visibility="collapsed")

    tab_map = {
        "① 프로젝트": tab_project,
        "② 대시보드": tab_dashboard,
        "③ 고용승계·계약서": tab_employees,
        "④ 인력구조 분석": tab_workforce,
        "⑤ 사규 비교": tab_regulation,
        "⑥ 직급 매핑": tab_grade_mapping,
        "⑦ 복지 매핑": tab_welfare_mapping,
        "⑧ 리스크": tab_risks,
        "⑨ 추진일정": tab_schedule,
        "⑩ 보고서": tab_report,
        "⑪ 자유자료 공유": tab_free_share,
        "⑫ 댓글/검토의견": tab_comments,
        "⑬ 팀 공유": tab_team_share,
        "⑭ 팀장 상황판": tab_manager_dashboard,
    }

    tab_map[menu]()


if __name__ == "__main__":
    main()
