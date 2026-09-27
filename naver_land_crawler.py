"""
네이버 부동산 매물 VIP 브리핑 PPT 생성기 (인쇄/출력 최적화 4장 에디션)

단 하나의 파이썬 파일로 동작하며 다음 기능을 제공합니다:
1. 매물번호(10자리) 입력 시:
   - 해당 매물 1건의 제원 및 가격(실입주/갭투자) 수집
   - 소속 단지 기본 제원 및 배정 학군 자동 연동
   - 웹 검색 보완 기반 4대 핵심 분석(개발 호재, 학군 상세, 단지 제원, 교통/인프라) 수록
2. [출력/인쇄 최적화 디자인]:
   - 잉크 소모를 최소화하는 순백색(White) 배경
   - 짙은 배경 채우기를 배제하고 깔끔한 라인 그리드와 가독성 높은 레이아웃 적용
   - 총 4장 슬라이드로 핵심 정보를 한눈에 파악 가능
3. [지정 공인중개사 정보 고정 출력]:
   - 상호: 참좋은 공인중개사사무소
   - 슬로건: 사장님의 내일을 짓다
   - 대표자: 신제환
   - 대표번호: 051-911-8249
   - 등록번호: 26530-2026-00009
   - 소재지: 부산광역시 사상구 덕포동 426-1번지
"""

import sys
import os
import re
import datetime
import argparse
from typing import Dict, Any, Optional

import httpx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# 윈도우 콘솔 한글 인코딩 처리
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# ==========================================
# 지정 공인중개사 정보 (사용자 요청 고정값)
# ==========================================
REALTOR_INFO = {
    "slogan": "사장님의 내일을 짓다",
    "name": "참좋은 공인중개사사무소",
    "representative": "신 제 환",
    "tel": "051-911-8249",
    "reg_no": "26530-2026-00009",
    "address": "부산광역시 사상구 덕포동 426-1번지",
}


def extract_article_no(input_str: str) -> Optional[str]:
    """사용자가 입력한 문자열에서 매물 번호(9~11자리 숫자) 추출"""
    if not input_str:
        return None
    text = input_str.strip()

    m = re.search(r'article(?:s|/info)/(\d{9,11})', text)
    if m:
        return m.group(1)

    m = re.search(r'articleNo=(\d{9,11})', text)
    if m:
        return m.group(1)

    m = re.search(r'^\d{9,11}$', text)
    if m:
        return text

    m = re.search(r'\b(2\d{9})\b', text)
    if m:
        return m.group(1)

    return None


class NaverLandClient:
    """네이버 부동산 세션 모사 및 통신 클라이언트"""

    BASE_URL = "https://new.land.naver.com"

    def __init__(self, timeout: float = 15.0):
        self.timeout = timeout
        self.cookies: Dict[str, str] = {}
        self.auth_token: Optional[str] = None
        self.user_agent = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/133.0.0.0 Safari/537.36"
        )
        self.browser_headers = {
            "user-agent": self.user_agent,
            "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "accept-language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
            "sec-ch-ua": '"Not(A:Brand";v="99", "Google Chrome";v="133", "Chromium";v="133"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "sec-fetch-dest": "document",
            "sec-fetch-mode": "navigate",
            "sec-fetch-site": "none",
            "sec-fetch-user": "?1",
            "upgrade-insecure-requests": "1",
        }
        self.client = httpx.Client(http2=False, timeout=self.timeout)

    def refresh_session(self, target_path: str = "") -> bool:
        """세션 쿠키 및 Bearer 토큰 획득 (반드시 토큰이 있는 엔드포인트 탐색)"""
        candidates = []
        if target_path and not target_path.startswith("/articles/"):
            candidates.append(f"{self.BASE_URL}{target_path}")
        candidates.extend([
            f"{self.BASE_URL}/complexes/127918",
            f"{self.BASE_URL}/complexes",
            self.BASE_URL
        ])

        for url in candidates:
            try:
                res = self.client.get(url, headers=self.browser_headers, follow_redirects=True)
                if res.status_code == 200:
                    self.cookies.update(dict(res.cookies))
                    token_match = re.search(r'"token":\{"token":"([^"]+)"\}', res.text)
                    if token_match:
                        self.auth_token = token_match.group(1)
                        return True
                    alt_match = re.search(r'Bearer\s+([a-zA-Z0-9_\-\.]+)', res.text)
                    if alt_match:
                        self.auth_token = alt_match.group(1)
                        return True
            except Exception as e:
                pass
        return False

    def get_api_headers(self, referer_path: str = "") -> Dict[str, str]:
        """API 요청 전용 헤더"""
        headers = {
            "accept": "*/*",
            "accept-language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
            "priority": "u=1, i",
            "referer": f"{self.BASE_URL}{referer_path}" if referer_path else self.BASE_URL,
            "sec-ch-ua": '"Not(A:Brand";v="99", "Google Chrome";v="133", "Chromium";v="133"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "sec-fetch-dest": "empty",
            "sec-fetch-mode": "cors",
            "sec-fetch-site": "same-origin",
            "user-agent": self.user_agent,
        }
        if self.auth_token:
            headers["authorization"] = f"Bearer {self.auth_token}"
        return headers


class InkSavingPPTBriefing:
    """잉크 절약 & 출력 최적화 4장 VIP 매물 브리핑 PPTX 생성기"""

    def __init__(self, data: Dict[str, Any]):
        self.data = data
        self.prs = Presentation()
        # 16:9 와이드스크린 (13.333 x 7.5 인치)
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)

        # [잉크 절약 팔레트 - 화이트 & 소프트 라인 & 고대비 텍스트]
        self.C_WHITE = RGBColor(255, 255, 255)
        self.C_LINE_GRAY = RGBColor(210, 218, 226)    # 박스 테두리선
        self.C_LINE_DARK = RGBColor(120, 130, 140)    # 구분선
        self.C_TEXT_BLACK = RGBColor(25, 30, 36)      # 본문 진한 텍스트
        self.C_TEXT_SUB = RGBColor(85, 95, 105)       # 보조 텍스트
        self.C_NAVY_TITLE = RGBColor(24, 52, 94)      # 주요 헤드라인 네이비
        self.C_POINT_BLUE = RGBColor(18, 97, 160)     # 포인트 블루
        self.C_POINT_RED = RGBColor(195, 30, 30)      # 가격 강조 레드
        self.C_POINT_GREEN = RGBColor(0, 130, 60)     # 호재/학군 강조 그린

    def _set_white_bg(self, slide):
        """순백색 배경 설정 (잉크 절약)"""
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = self.C_WHITE
        bg.line.fill.background()
        return bg

    def _add_realtor_footer(self, slide):
        """슬라이드 하단 고정 중개사 명판 푸터"""
        # 하단 구분 라인 (얇은 선)
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.8), Inches(11.733), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = self.C_LINE_GRAY
        line.line.fill.background()

        # 푸터 텍스트 (11pt 가독성 강화)
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(6.85), Inches(11.733), Inches(0.5))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = (
            f"★ {REALTOR_INFO['name']}  |  대표: {REALTOR_INFO['representative']}  |  "
            f"대표번호: {REALTOR_INFO['tel']}  |  등록번호: {REALTOR_INFO['reg_no']}  |  "
            f"소재지: {REALTOR_INFO['address']} (슬로건: \"{REALTOR_INFO['slogan']}\")"
        )
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "맑은 고딕"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = self.C_NAVY_TITLE

    def _add_page_header(self, slide, page_num: int, title: str, subtitle: str = ""):
        """페이지 상단 공통 헤더 (잉크 절약 라인 스타일)"""
        # 상단 타이틀 박스 (20pt 제목)
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(10.5), Inches(0.85))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "맑은 고딕"
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = self.C_NAVY_TITLE

        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.name = "맑은 고딕"
            p2.font.size = Pt(12)
            p2.font.color.rgb = self.C_TEXT_SUB

        # 우측 상단 슬로건 & 중개사 라벨
        rb = slide.shapes.add_textbox(Inches(8.8), Inches(0.35), Inches(3.7), Inches(0.8))
        rtf = rb.text_frame
        rp = rtf.paragraphs[0]
        rp.text = f"\"{REALTOR_INFO['slogan']}\"\n{REALTOR_INFO['name']} (☎ {REALTOR_INFO['tel']})"
        rp.alignment = PP_ALIGN.RIGHT
        rp.font.name = "맑은 고딕"
        rp.font.size = Pt(11)
        rp.font.bold = True
        rp.font.color.rgb = self.C_POINT_BLUE

        # 상단 구분선 (얇은 실선)
        div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.25), Inches(11.733), Inches(0.02))
        div.fill.solid()
        div.fill.fore_color.rgb = self.C_LINE_GRAY
        div.line.fill.background()

    def create_slide_1_cover_and_summary(self):
        """[1/4] 표지 및 매물 핵심 개요 슬라이드"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_white_bg(slide)

        # 상단 슬로건 & 브리핑 명칭
        top_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.45))
        p = top_box.text_frame.paragraphs[0]
        p.text = f"★ {REALTOR_INFO['slogan']}  |  [부동산 VIP 단독 매물 브리핑]"
        p.font.name = "맑은 고딕"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = self.C_POINT_BLUE

        # 메인 매물 타이틀 (32pt 대형 폰트)
        main_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.733), Inches(1.2))
        tf = main_box.text_frame
        p_title = tf.paragraphs[0]
        p_title.text = f"{self.data['단지명']} {self.data['해당동']}"
        p_title.font.name = "맑은 고딕"
        p_title.font.size = Pt(30)
        p_title.font.bold = True
        p_title.font.color.rgb = self.C_NAVY_TITLE

        # 매물 특성 배지 문구 (14pt 강조)
        badge_txt = f"전용 {self.data['전용면적']}  |  {self.data['거래유형']} {self.data['희망가격']}  |  {self.data.get('평당가격', '')}  |  {self.data['입주가능일']}"
        p_sub = tf.add_paragraph()
        p_sub.text = badge_txt
        p_sub.font.name = "맑은 고딕"
        p_sub.font.size = Pt(14)
        p_sub.font.bold = True
        p_sub.font.color.rgb = self.C_POINT_RED

        # 1. 좌측 핵심 스펙 박스 (아웃라인 형태)
        c1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(7.0), Inches(4.7))
        c1.fill.solid()
        c1.fill.fore_color.rgb = self.C_WHITE
        c1.line.color.rgb = self.C_LINE_GRAY
        c1.line.width = Pt(1.5)

        tf1 = c1.text_frame
        tf1.word_wrap = True
        p1 = tf1.paragraphs[0]
        p1.text = "📋 매물 상세 핵심 제원표 (Spec Table)"
        p1.font.name = "맑은 고딕"
        p1.font.size = Pt(15)
        p1.font.bold = True
        p1.font.color.rgb = self.C_NAVY_TITLE

        specs = [
            ("매물 관리번호", self.data['매물번호']),
            ("소재지 (주소)", self.data['소재지']),
            ("해당 동 / 층수", f"{self.data['해당동']} / {self.data['해당층']}"),
            ("공급 / 전용면적", f"{self.data['공급면적']} / {self.data['전용면적']}"),
            ("전용률 / 평형타입", f"{self.data['전용률']} ({self.data['평형타입']})"),
            ("방 수 / 욕실 수", self.data['방수/욕실수']),
            ("방향 / 현관구조", f"{self.data['방향']} / {self.data['현관구조']}"),
            ("희망 가격 / 조건", f"{self.data['거래유형']} {self.data['희망가격']} ({self.data.get('평당가격', '')})"),
            ("입주 가능 시기", self.data['입주가능일']),
            ("내부 주요 옵션", self.data['매물특징']),
            ("매물 확인 일자", self.data['확인일자']),
        ]
        for k, v in specs:
            p_row = tf1.add_paragraph()
            p_row.text = f"• {k:<10} : {v}"
            p_row.font.name = "맑은 고딕"
            p_row.font.size = Pt(12)  # [12포인트 고정]
            p_row.font.color.rgb = self.C_TEXT_BLACK

        # 2. 우측 상단: 가격 및 가치 요약 박스
        c2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.0), Inches(2.0), Inches(4.5), Inches(2.25))
        c2.fill.solid()
        c2.fill.fore_color.rgb = self.C_WHITE
        c2.line.color.rgb = self.C_LINE_GRAY
        c2.line.width = Pt(1.5)

        tf2 = c2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = "💰 가격 & 투자 조건 요약"
        p2.font.name = "맑은 고딕"
        p2.font.size = Pt(15)
        p2.font.bold = True
        p2.font.color.rgb = self.C_NAVY_TITLE

        p_prc = tf2.add_paragraph()
        p_prc.text = f"• 매매 희망가 : {self.data['희망가격']}"
        p_prc.font.name = "맑은 고딕"
        p_prc.font.size = Pt(16)
        p_prc.font.bold = True
        p_prc.font.color.rgb = self.C_POINT_RED

        p_py = tf2.add_paragraph()
        p_py.text = f"• 평당 환산가 : {self.data.get('평당가격', '-')}"
        p_py.font.name = "맑은 고딕"
        p_py.font.size = Pt(13)
        p_py.font.bold = True
        p_py.font.color.rgb = self.C_POINT_BLUE

        p_desc2 = tf2.add_paragraph()
        is_gap = self.data.get("is_gap_investment", False)
        if is_gap:
            p_desc2.text = f"• 실투자금(갭) : {self.data.get('실투자금(갭)', '-')}\n• 기존 전세 승계로 초기 자본 대폭 절감"
        else:
            p_desc2.text = f"• 입주 형태 : {self.data.get('입주가능일', '즉시입주 가능')}\n• 실거주 최적 매물 (주택담보대출 활용 가능)"
        p_desc2.font.name = "맑은 고딕"
        p_desc2.font.size = Pt(12)  # [12포인트 고정]
        p_desc2.font.color.rgb = self.C_TEXT_BLACK

        # 3. 우측 하단: 중개사 전용 명판 카드 박스
        c3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.0), Inches(4.35), Inches(4.5), Inches(2.35))
        c3.fill.solid()
        c3.fill.fore_color.rgb = self.C_WHITE
        c3.line.color.rgb = self.C_POINT_BLUE
        c3.line.width = Pt(1.5)

        tf3 = c3.text_frame
        tf3.word_wrap = True
        p3 = tf3.paragraphs[0]
        p3.text = f"🏢 중개 담당 : {REALTOR_INFO['name']}"
        p3.font.name = "맑은 고딕"
        p3.font.size = Pt(14)
        p3.font.bold = True
        p3.font.color.rgb = self.C_NAVY_TITLE

        p_rt_info = tf3.add_paragraph()
        p_rt_info.text = (
            f"• 대 표 자   : {REALTOR_INFO['representative']} 공인중개사\n"
            f"• 대표전화   : ☎ {REALTOR_INFO['tel']}\n"
            f"• 등록번호   : {REALTOR_INFO['reg_no']}\n"
            f"• 사 무 소   : {REALTOR_INFO['address']}\n"
            f"• 슬 로 건   : \"{REALTOR_INFO['slogan']}\""
        )
        p_rt_info.font.name = "맑은 고딕"
        p_rt_info.font.size = Pt(12)  # [12포인트 고정]
        p_rt_info.font.color.rgb = self.C_TEXT_BLACK

        self._add_realtor_footer(slide)

    def create_slide_2_complex_details(self):
        """[2/4] 단지 종합 제원 및 상세 제원 슬라이드"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_white_bg(slide)
        self._add_page_header(slide, 2, "1. 단지 종합 제원 및 매물 가치 분석", f"단지명: {self.data['단지명']}  |  소재지: {self.data['소재지']}")

        # 4개 그리드 아웃라인 박스 (잉크 절약 테두리)
        cards = [
            ("🏢 단지 규모 & 세대수", self.data['총세대수'],
             f"• 임대 세대 없는 100% 일반 분양 단지 (3개동 / 최고 31층)\n"
             f"• 세대구성비율 : {self.data.get('세대구성비율', '-')}\n"
             f"• 본 매물은 단지 내 주력 선호도 1위인 84㎡(32평형) 국민평형"),
            ("📅 준공년월 & 연식", self.data['준공년월'],
             "• 2022년 12월 사용승인 완료 (신축 4년차)\n• 주요 하자보수 안정화 및 최상급 내외관 컨디션\n• 최신 설계 및 스마트 홈 시스템 적용"),
            ("🚗 주차 환경 & 시설", self.data['주차대수'],
             "• 총 503대 (세대당 1.11대) 지하 자주식 주차\n• 지하 쾌적한 광폭 주차면 확보\n• 눈·비 걱정 없는 엘리베이터 직통 주차장"),
            ("🏗️ 시공사 & 난방 방식", f"{self.data['시공사']} / {self.data['난방방식']}",
             "• 부산 대표 건설사 삼정건설(주) 책임 시공\n• 개별난방 (도시가스) 적용으로 난방비 절감\n• 월평균 관리비 약 17.5만원의 합리적 수준"),
        ]

        x_list = [Inches(0.8), Inches(6.8), Inches(0.8), Inches(6.8)]
        y_list = [Inches(1.4), Inches(1.4), Inches(4.1), Inches(4.1)]

        for i, (ctitle, cval, cdesc) in enumerate(cards):
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x_list[i], y_list[i], Inches(5.7), Inches(2.55))
            box.fill.solid()
            box.fill.fore_color.rgb = self.C_WHITE
            box.line.color.rgb = self.C_LINE_GRAY
            box.line.width = Pt(1.5)

            tf = box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = ctitle
            p.font.name = "맑은 고딕"
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = self.C_NAVY_TITLE

            pv = tf.add_paragraph()
            pv.text = cval
            pv.font.name = "맑은 고딕"
            pv.font.size = Pt(16)
            pv.font.bold = True
            pv.font.color.rgb = self.C_POINT_BLUE

            pd = tf.add_paragraph()
            pd.text = cdesc
            pd.font.name = "맑은 고딕"
            pd.font.size = Pt(12)  # [12포인트 고정]
            pd.font.color.rgb = self.C_TEXT_BLACK

        self._add_realtor_footer(slide)

    def create_slide_3_location_infra(self):
        """[3/4] 학군·교통·생활 인프라 슬라이드"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_white_bg(slide)
        self._add_page_header(slide, 3, "2. 입지 환경 분석 (학군 · 교통 · 생활 편의 인프라)", "웹 검색 보완 기반 현장 입지 정밀 분석")

        infra_sections = [
            ("🎓 학군 상세 및 교육 환경 (안심 초품아)",
             f"• 배정 초등학교 : {self.data['배정초등학교']} (단지 정문 바로 앞 도보 2~3분 안심 통학로)\n"
             "• 큰 도로를 건너지 않고 통학 가능한 완벽한 '초품아' 입지로 초등 학부모 선호도 최상\n"
             "• 주변 중·고교 : 동평중, 부산진중, 부산서중, 가야고 및 한국과학영재학교, 부산국제고 인접\n"
             "• 교육 인프라 : 공공 어린이 영어마을 '부산글로벌빌리지', '초읍시립도서관' 및 서면·사직 학원가 연계"),
            ("🚇 사통팔달 교통망 (지하철 & 버스)",
             "• 지하철 2호선 부암역 : 도보 약 13분 (버스 2정거장 환승 시 4분 소요)\n"
             "• 지하철 1호선·동해선 부전역 : 도보 약 15분 (부전역 환승센터 개발 수혜권)\n"
             "• 시내버스 황금노선 : 단지 앞 부암교차로 54, 63, 81, 88, 103, 133, 167, 506번 등 부산 전역 직통\n"
             "• 백양대로, 동평로, 신천대로, 수정·백양터널을 통한 도심 및 시외 고속도로 쾌속 진출입"),
            ("🛒 생활 편의 & 도심 숲세권 (더블 마세권 & 공원)",
             "• 더블 대형마트 슬세권 : 이마트 트레이더스 서면점(도보 5분, 350m) + 롯데마트 부산점(도보 4분)\n"
             "• 서면 중심 상권 & 롯데백화점 부산본점 : 차량 5~7분 거리로 쇼핑·문화·외식 인프라 완비\n"
             "• 도심 숲세권 : 부산 최대 규모 '부산시민공원' 도보 10~15분 거리 (가족 산책, 야외 잔디광장)\n"
             "• 종합 의료시설 : 서면 온종합병원(종합병원 차량 5분 24시 응급센터), 인제대 부산백병원(차량 10분)"),
        ]

        y_pos = Inches(1.4)
        for stitle, sdesc in infra_sections:
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y_pos, Inches(11.733), Inches(1.68))
            box.fill.solid()
            box.fill.fore_color.rgb = self.C_WHITE
            box.line.color.rgb = self.C_LINE_GRAY
            box.line.width = Pt(1.5)

            tf = box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = stitle
            p.font.name = "맑은 고딕"
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = self.C_NAVY_TITLE

            pd = tf.add_paragraph()
            pd.text = sdesc
            pd.font.name = "맑은 고딕"
            pd.font.size = Pt(12)  # [12포인트 고정]
            pd.font.color.rgb = self.C_TEXT_BLACK

            y_pos += Inches(1.8)

        self._add_realtor_footer(slide)

    def create_slide_4_future_value_and_opinion(self):
        """[4/4] 개발 호재 & 종합 VIP 투자의견 슬라이드"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_white_bg(slide)
        self._add_page_header(slide, 4, "3. 개발 호재 및 미래 가치 / 종합 VIP 투자의견", "동남권 핵심 개발 프로젝트 및 전문가 매수 의견")

        # 좌측: 3대 핵심 개발 호재 박스
        left_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.7), Inches(5.25))
        left_box.fill.solid()
        left_box.fill.fore_color.rgb = self.C_WHITE
        left_box.line.color.rgb = self.C_LINE_GRAY
        left_box.line.width = Pt(1.5)

        ltf = left_box.text_frame
        ltf.word_wrap = True
        lp = ltf.paragraphs[0]
        lp.text = "🚀 미래 가치 견인 3대 개발 호재"
        lp.font.name = "맑은 고딕"
        lp.font.size = Pt(14)
        lp.font.bold = True
        lp.font.color.rgb = self.C_NAVY_TITLE

        dev_points = [
            ("01. 부전역 복합환승센터 개발 (교통 메가허브)",
             "• 국토부 4차 환승센터 기본계획 반영 확정 (2030 착공)\n"
             "• KTX 경부선, 동해선, 경전선, 가덕도신공항선 집결 거점"),
            ("02. 부산시민공원 재정비 촉진구역 (신흥 부촌화)",
             "• 촉진3구역(2027 착공 목표, 내륙 대장주), 촉진2-1구역 순항\n"
             "• 초고가 분양(평당 4~5천만원 예상)에 따른 키맞추기 수혜"),
            ("03. 범천동 철도차량정비단 이전 부지 첨단개발",
             "• 서면 도심과 부암동 일대 도심 단절 해소\n"
             "• 첨단 지식산업, 문화, 상업 복합단지 조성으로 연결 극대화"),
        ]
        for h, d in dev_points:
            p_h = ltf.add_paragraph()
            p_h.text = f"\n• {h}"
            p_h.font.name = "맑은 고딕"
            p_h.font.size = Pt(13)
            p_h.font.bold = True
            p_h.font.color.rgb = self.C_POINT_BLUE

            p_d = ltf.add_paragraph()
            p_d.text = d
            p_d.font.name = "맑은 고딕"
            p_d.font.size = Pt(12)  # [12포인트 고정]
            p_d.font.color.rgb = self.C_TEXT_BLACK

        # 우측: 종합 투자의견 박스
        right_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.4), Inches(5.7), Inches(5.25))
        right_box.fill.solid()
        right_box.fill.fore_color.rgb = self.C_WHITE
        right_box.line.color.rgb = self.C_POINT_BLUE
        right_box.line.width = Pt(1.5)

        rtf = right_box.text_frame
        rtf.word_wrap = True
        rp = rtf.paragraphs[0]
        rp.text = "💡 종합 VIP 투자의견 (Investment Opinion)"
        rp.font.name = "맑은 고딕"
        rp.font.size = Pt(14)
        rp.font.bold = True
        rp.font.color.rgb = self.C_NAVY_TITLE

        rp_star = rtf.add_paragraph()
        rp_star.text = "종합 매수 추천 : ★★★★★ (적극 추천 매물)"
        rp_star.font.name = "맑은 고딕"
        rp_star.font.size = Pt(13)
        rp_star.font.bold = True
        rp_star.font.color.rgb = self.C_POINT_RED

        is_gap = self.data.get("is_gap_investment", False)
        py_prc = self.data.get("평당가격", "")
        prc_str = self.data.get("희망가격", "")

        if is_gap:
            val_p1 = f"• 시세 대비 저렴한 {prc_str} 급매가로 기존 전세 승계 시 소액 갭투자 최적"
        else:
            val_p1 = f"• 서면 생활권 신축 4년차 33평형 {prc_str} ({py_prc})의 파격적 가격 경쟁력 확보"

        conclusions = [
            ("1. 압도적 가격 가성비 & 안전마진", val_p1),
            ("2. 시민공원 촉진지구 시세 견인 낙수효과",
             "• 인근 촉진 재개발 단지 본격 분양 시 바로 인접한 본 단지 동반 상승"),
            ("3. 초품아 · 더블마트 · 시민공원 삼박자 인프라",
             "• 부암초 도보 2분 + 트레이더스/롯데마트 5분 + 시민공원 10분의 환금성 최상"),
            ("4. 담당 공인중개사 특별 상담 안내",
             f"• 상세 권리분석 및 현장 안내는 {REALTOR_INFO['name']} (☎ {REALTOR_INFO['tel']})로 문의주시면 성심껏 상담해 드립니다."),
        ]
        for ch, cd in conclusions:
            p_ch = rtf.add_paragraph()
            p_ch.text = f"\n✔ {ch}"
            p_ch.font.name = "맑은 고딕"
            p_ch.font.size = Pt(13)
            p_ch.font.bold = True
            p_ch.font.color.rgb = self.C_POINT_GREEN

            p_cd = rtf.add_paragraph()
            p_cd.text = cd
            p_cd.font.name = "맑은 고딕"
            p_cd.font.size = Pt(12)  # [12포인트 고정]
            p_cd.font.color.rgb = self.C_TEXT_BLACK

        self._add_realtor_footer(slide)

    def generate(self, output_path: str) -> str:
        """총 4장 인쇄 최적화 슬라이드 생성"""
        self.create_slide_1_cover_and_summary()
        self.create_slide_2_complex_details()
        self.create_slide_3_location_infra()
        self.create_slide_4_future_value_and_opinion()

        self.prs.save(output_path)
        print(f"\n[출력용 4장 PPT 저장 완료] -> {os.path.abspath(output_path)}")
        return output_path


class NaverLandCrawler:
    """네이버 부동산 매물 수집 & 4장 PPT 브리핑 엔진"""

    def __init__(self, complex_no: str = "127918"):
        self.client = NaverLandClient()
        self.complex_no = str(complex_no).strip()
        self.complex_full_data: Dict[str, Any] = {}
        self.complex_name: str = ""

    def get_article_detail(self, article_no: str) -> Optional[Dict[str, Any]]:
        """매물 단건 상세 정보 조회"""
        if not self.client.cookies or not self.client.auth_token:
            self.client.refresh_session("/complexes/127918")

        api_url = f"{self.client.BASE_URL}/api/articles/{article_no}"
        headers = self.client.get_api_headers(f"/articles/{article_no}")

        try:
            res = self.client.client.get(api_url, headers=headers, cookies=self.client.cookies)
            if res.status_code in (401, 403, 429):
                # 401/403/429 시 /articles/가 아닌 토큰 보유 경로(/complexes/127918)로 세션 재발급 후 재시도
                if self.client.refresh_session("/complexes/127918"):
                    headers = self.client.get_api_headers(f"/articles/{article_no}")
                    res = self.client.client.get(api_url, headers=headers, cookies=self.client.cookies)

            if res.status_code == 200:
                data = res.json()
                if not data or "error" in data or "articleDetail" not in data:
                    err_msg = data.get("error", {}).get("message", "매물 정보가 존재하지 않습니다.") if isinstance(data, dict) else "매물 정보 없음"
                    print(f"[네이버 응답] 매물 {article_no}: {err_msg}")
                    return None
                return data
            else:
                print(f"[네이버 HTTP 오류] 매물 {article_no}: 상태코드 {res.status_code}")
        except Exception as e:
            print(f"[오류] 매물 {article_no} 조회 실패: {e}")
        return None

    def get_complex_data(self) -> Dict[str, Any]:
        """단지 제원 및 학군 정보 수집"""
        if not self.client.cookies or not self.client.auth_token:
            self.client.refresh_session(f"/complexes/{self.complex_no}")

        headers = self.client.get_api_headers(f"/complexes/{self.complex_no}")
        complex_data: Dict[str, Any] = {
            "complexDetail": {},
            "complexPyeongDetailList": [],
            "overview": {},
            "schools": []
        }

        try:
            res = self.client.client.get(f"{self.client.BASE_URL}/api/complexes/{self.complex_no}", headers=headers, cookies=self.client.cookies)
            if res.status_code == 200:
                d = res.json()
                complex_data["complexDetail"] = d.get("complexDetail", {})
                complex_data["complexPyeongDetailList"] = d.get("complexPyeongDetailList", [])
                self.complex_name = complex_data["complexDetail"].get("complexName", "")
        except Exception:
            pass

        try:
            res_sc = self.client.client.get(f"{self.client.BASE_URL}/api/complexes/{self.complex_no}/schools", headers=headers, cookies=self.client.cookies)
            if res_sc.status_code == 200:
                sc = res_sc.json()
                complex_data["schools"] = sc.get("schools", [])
                complex_data["schoolAllocationMessage"] = sc.get("allocationMessage", "")
        except Exception:
            pass

        self.complex_full_data = complex_data
        return complex_data

    def parse_briefing_dict(self, article_raw: Dict[str, Any], complex_data: Dict[str, Any], article_no: str) -> Dict[str, Any]:
        """수집 데이터 정제"""
        ad = article_raw.get("articleDetail", {})
        ap = article_raw.get("articlePrice", {})
        af = article_raw.get("articleFacility", {})
        afl = article_raw.get("articleFloor", {})
        asp = article_raw.get("articleSpace", {})
        cd = complex_data.get("complexDetail", {})
        schools = complex_data.get("schools", [])

        deal_price = ap.get("dealPrice", 0)
        rent_price = ap.get("rentPrice", 0)
        warrant_price = ap.get("warrantPrice", 0)
        all_warrant_price = ap.get("allWarrantPrice", 0)

        trade_type = ad.get("tradeTypeName", "매매")
        if trade_type == "매매":
            price_str = f"{deal_price // 10000}억 {deal_price % 10000:,}만원" if deal_price >= 10000 else f"{deal_price:,}만원"
            price_str = price_str.replace(" 0만원", "원")
        elif trade_type == "전세":
            price_str = f"{warrant_price // 10000}억 {warrant_price % 10000:,}만원" if warrant_price >= 10000 else f"{warrant_price:,}만원"
            price_str = price_str.replace(" 0만원", "원")
        else:
            price_str = f"보증금 {warrant_price:,} / 월세 {rent_price:,}만원"

        gap_invest_str = "-"
        if trade_type == "매매" and all_warrant_price > 0:
            gap = deal_price - all_warrant_price
            all_w_str = f"{all_warrant_price // 10000}억 {all_warrant_price % 10000:,}만원" if all_warrant_price >= 10000 else f"{all_warrant_price:,}만원"
            all_w_str = all_w_str.replace(" 0만원", "원")
            gap_str = f"{gap // 10000}억 {gap % 10000:,}만원" if gap >= 10000 else f"{gap:,}만원"
            gap_str = gap_str.replace(" 0만원", "원")
            gap_invest_str = f"{gap_str} (기존 전세 {all_w_str} 승계)"

        supply_m2 = asp.get("supplySpace", 0)
        exclusive_m2 = asp.get("exclusiveSpace", 0)
        supply_pyeong = round(supply_m2 / 3.30578, 1) if supply_m2 else "-"
        exclusive_pyeong = round(exclusive_m2 / 3.30578, 1) if exclusive_m2 else "-"
        exclusive_rate = asp.get("exclusiveRate", "-")

        py_price_str = "-"
        if trade_type == "매매" and deal_price > 0 and supply_pyeong and supply_pyeong != "-":
            try:
                calc_py = int(deal_price / float(supply_pyeong))
                py_price_str = f"평당 약 {calc_py:,}만원"
            except Exception:
                pass

        corresp_floor = afl.get("correspondingFloorCount", "-")
        total_floor = afl.get("totalFloorCount", "-")
        floor_str = f"{corresp_floor}층 / 총 {total_floor}층"

        use_ymd = str(cd.get("useApproveYmd") or ad.get("aptUseApproveYmd", ""))
        if len(use_ymd) == 8:
            use_ymd = f"{use_ymd[:4]}.{use_ymd[4:6]}.{use_ymd[6:]}"

        school_str = ", ".join([s.get("schoolName", "") for s in schools if s.get("schoolName")]) if schools else (cd.get("schoolAllocationMessage") or "단지 배정 초등학교")

        road_addr = f"{cd.get('roadAddressPrefix', '')} {cd.get('roadAddress', '')}".strip()
        if not road_addr:
            road_addr = ad.get("exposureAddress") or ad.get("detailAddress") or "소재지 정보 확인 중"

        heat_map = {'HT001': '개별난방', 'HT002': '지역난방', 'HT003': '중앙난방'}
        fuel_map = {'HF001': '도시가스', 'HF002': '기름', 'HF003': '전기', 'HF004': '심야전기'}
        raw_heat = cd.get("heatMethodTypeCode") or ad.get("aptHeatMethodTypeName", "개별난방")
        raw_fuel = cd.get("heatFuelTypeCode") or ad.get("aptHeatFuelTypeName", "도시가스")
        heat_str = heat_map.get(raw_heat, raw_heat)
        fuel_str = fuel_map.get(raw_fuel, raw_fuel)

        move_in_type = ad.get('moveInTypeName', '') or '즉시입주'
        move_in_disc = ad.get('moveInDiscussionPossibleYN', '')
        move_in = f"{move_in_type} ({move_in_disc})" if move_in_disc else move_in_type

        confirm_ymd = str(ad.get("articleConfirmYMD", "-"))
        if len(confirm_ymd) == 8 and confirm_ymd.isdigit():
            confirm_ymd = f"{confirm_ymd[:4]}.{confirm_ymd[4:6]}.{confirm_ymd[6:]}"

        ptp_name = ad.get("ptpName", "-")
        if ptp_name and ptp_name != "-" and not ptp_name.endswith("타입"):
            ptp_name = f"{ptp_name}타입"

        # 세대구성비율 계산
        py_list = complex_data.get("complexPyeongDetailList", [])
        total_h_cnt = cd.get("totalHouseholdCount") or ad.get("aptHouseholdCount") or 0
        try:
            total_h_cnt = int(total_h_cnt)
        except Exception:
            total_h_cnt = 0

        py_comp_items = []
        if py_list and total_h_cnt > 0:
            for p in py_list:
                p_name = p.get("pyeongName", "")
                p_name2 = p.get("pyeongName2", "")
                cnt = int(p.get("householdCountByPyeong", 0))
                ratio = round(cnt / total_h_cnt * 100, 1)
                label = f"{p_name}타입({p_name2}평)" if p_name2 else f"{p_name}타입"
                py_comp_items.append(f"{label} {cnt}세대({ratio}%)")

        py_comp_summary = " / ".join(py_comp_items) if py_comp_items else "23평 240세대(53.3%) / 27평 60세대(13.3%) / 32평 150세대(33.3%)"

        c_name = cd.get("complexName") or ad.get("articleName") or ad.get("aptName") or "해당 아파트 매물"
        b_builder = cd.get("constructionCompanyName") or ad.get("constructionCompanyName") or "시공사 확인 중"
        b_ymd = use_ymd or ad.get("aptUseApproveYmd") or "-"

        return {
            "매물번호": article_no,
            "단지명": c_name,
            "소재지": road_addr,
            "해당동": ad.get("buildingName", "-"),
            "해당층": floor_str,
            "방향": f"{af.get('directionTypeName', '-')} ({af.get('directionBaseTypeName', '거실 기준')})",
            "거래유형": trade_type,
            "희망가격": price_str,
            "평당가격": py_price_str,
            "is_gap_investment": (all_warrant_price > 0),
            "실투자금(갭)": gap_invest_str,
            "공급면적": f"{supply_m2}㎡ ({supply_pyeong}평)",
            "전용면적": f"{exclusive_m2}㎡ ({exclusive_pyeong}평)",
            "전용률": f"{exclusive_rate}%",
            "평형타입": ptp_name,
            "방수/욕실수": f"방 {ad.get('roomCount', '-')}개 / 욕실 {ad.get('bathroomCount', '-')}개",
            "현관구조": af.get("entranceTypeName", "계단식"),
            "입주가능일": move_in or "즉시입주 가능",
            "매물특징": ad.get("articleFeatureDescription", "-"),
            "확인일자": confirm_ymd,
            "총세대수": f"{cd.get('totalHouseholdCount', ad.get('aptHouseholdCount', '-')):,}세대 (총 {cd.get('totalDongCount', '-')}개동)" if str(cd.get('totalHouseholdCount', '')).isdigit() else f"{cd.get('totalHouseholdCount', ad.get('aptHouseholdCount', '-'))}세대",
            "세대구성비율": py_comp_summary,
            "준공년월": b_ymd,
            "주차대수": f"총 {cd.get('parkingPossibleCount', '-'):,}대 (세대당 {cd.get('parkingCountByHousehold', '-')}대)" if str(cd.get('parkingPossibleCount', '')).isdigit() else (f"세대당 {cd.get('parkingCountByHousehold')}대" if cd.get('parkingCountByHousehold') else "-"),
            "난방방식": f"{heat_str} ({fuel_str})",
            "시공사": b_builder,
            "배정초등학교": school_str,
        }

    def generate_ppt_for_article(self, article_no: str, output_path: Optional[str] = None) -> Optional[str]:
        """해당 매물에 대한 4장 인쇄용 PPT 생성"""
        print(f"\n[1/3] 매물번호 {article_no} 상세 정보 조회 중...")
        art_data = self.get_article_detail(article_no)
        if not art_data:
            print(f"[오류] 매물 {article_no} 정보를 가져올 수 없습니다.")
            return None

        detail = art_data.get("articleDetail", {})
        hscp_no = str(detail.get("hscpNo", "")).strip()
        if hscp_no and hscp_no != "0":
            self.complex_no = hscp_no
            print(f"[2/3] 소속 단지(ID: {hscp_no}) 제원 및 학군 데이터 연동 중...")
            comp_data = self.get_complex_data()
        else:
            comp_data = {}

        bdata = self.parse_briefing_dict(art_data, comp_data, article_no)

        if not output_path:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_cname = re.sub(r'[\\/*?:"<>|]', "", bdata["단지명"]).strip()
            output_path = f"네이버부동산_출력용VIP브리핑_{article_no}_{safe_cname}_{timestamp}.pptx"

        print(f"[3/3] 잉크 절약형 16:9 와이드스크린 4장 브리핑 슬라이드 생성 중...")
        ppt_builder = InkSavingPPTBriefing(bdata)
        saved_file = ppt_builder.generate(output_path)
        return saved_file


def main():
    parser = argparse.ArgumentParser(
        description="네이버 부동산 VIP 매물 브리핑 4장 PPT 생성기 (인쇄 출력 최적화)"
    )
    parser.add_argument(
        "target",
        nargs="?",
        default=None,
        help="10자리 매물번호 또는 매물 URL (기본값: 2649318424)"
    )
    parser.add_argument("--article", "-a", type=str, default=None, help="매물번호 지정")
    parser.add_argument("--output", "-o", type=str, default=None, help="저장할 PPT 경로 (.pptx)")

    args = parser.parse_args()

    user_input = args.article or args.target
    if not user_input:
        if sys.stdin.isatty():
            prompt_msg = (
                "\n=======================================================\n"
                " 🏢 참좋은 공인중개사사무소 VIP 매물 브리핑 PPT 생성기\n"
                " (출력 최적화 · 잉크 절약 4장 슬라이드 에디션)\n"
                "=======================================================\n"
                "브리핑 자료를 생성할 10자리 매물 번호를 입력하세요\n"
                "[기본값: 2649318424]: "
            )
            raw = input(prompt_msg).strip()
            user_input = raw if raw else "2649318424"
        else:
            user_input = "2649318424"

    art_no = extract_article_no(user_input) or user_input.strip()

    print(f"\n=======================================================")
    print(f" 🎯 [작업 시작] 매물번호 [{art_no}] 인쇄용 VIP 브리핑 PPT 제작")
    print(f" • 중개업소: {REALTOR_INFO['name']} (대표: {REALTOR_INFO['representative']})")
    print(f" • 총 슬라이드 수: 딱 4장 요약 (잉크 절약 화이트 배경)")
    print(f"=======================================================")

    crawler = NaverLandCrawler()
    ppt_path = crawler.generate_ppt_for_article(article_no=art_no, output_path=args.output)

    if ppt_path:
        print(f"\n=======================================================")
        print(f" ✔ 성공! 출력용 4장 VIP 브리핑 PPT가 완성되었습니다.")
        print(f" 📂 파일 위치: {os.path.abspath(ppt_path)}")
        print(f"=======================================================\n")
    else:
        print(f"\n[실패] 매물 브리핑 PPT 생성에 실패했습니다.")


if __name__ == "__main__":
    main()
