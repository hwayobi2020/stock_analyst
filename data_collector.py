"""
종목 데이터 수집: yfinance (재무) + Google News RSS (뉴스)
"""
import json
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
import yfinance as yf


def is_korean_stock(query):
    """한국 종목인지 판별"""
    # 숫자로만 구성되면 한국 종목코드
    if query.strip().isdigit():
        return True
    # 한글이 포함되면 한국 종목
    for ch in query:
        if "\uac00" <= ch <= "\ud7a3":
            return True
    return False


# 주요 한국 종목 매핑 (자주 검색되는 종목)
KOREAN_STOCKS = {
    "삼성전자": "005930.KS",
    "SK하이닉스": "000660.KS",
    "LG에너지솔루션": "373220.KS",
    "삼성바이오로직스": "207940.KS",
    "현대차": "005380.KS",
    "현대자동차": "005380.KS",
    "기아": "000270.KS",
    "셀트리온": "068270.KS",
    "KB금융": "105560.KS",
    "신한지주": "055550.KS",
    "POSCO홀딩스": "005490.KS",
    "포스코홀딩스": "005490.KS",
    "네이버": "035420.KS",
    "카카오": "035720.KS",
    "LG화학": "051910.KS",
    "삼성SDI": "006400.KS",
    "현대모비스": "012330.KS",
    "SK이노베이션": "096770.KS",
    "삼성물산": "028260.KS",
    "LG전자": "066570.KS",
    "한국전력": "015760.KS",
    "SK텔레콤": "017670.KS",
    "KT": "030200.KS",
    "하나금융지주": "086790.KS",
    "우리금융지주": "316140.KS",
    "한화에어로스페이스": "012450.KS",
    "두산에너빌리티": "034020.KS",
    "크래프톤": "259960.KS",
    "카카오뱅크": "323410.KS",
    "엔씨소프트": "036570.KS",
}


def resolve_ticker(query):
    """종목명을 yfinance 티커로 변환"""
    query = query.strip()

    # 이미 티커 형식이면 그대로 반환
    if "." in query or query.isupper():
        return query

    # 한국 종목 매핑에서 찾기
    if query in KOREAN_STOCKS:
        return KOREAN_STOCKS[query]

    # 숫자 코드면 .KS 붙이기
    if query.isdigit():
        return f"{query}.KS"

    # 그 외는 그대로 (미국 종목 티커로 간주)
    return query.upper()


def get_stock_data(query):
    """종목 재무 데이터 수집"""
    ticker = resolve_ticker(query)
    t = yf.Ticker(ticker)
    info = t.info

    # 핵심 재무지표 추출
    data = {
        "ticker": ticker,
        "name": info.get("shortName") or info.get("longName") or query,
        "current_price": info.get("currentPrice") or info.get("regularMarketPrice"),
        "market_cap": info.get("marketCap"),
        "per": info.get("trailingPE"),
        "forward_per": info.get("forwardPE"),
        "pbr": info.get("priceToBook"),
        "roe": info.get("returnOnEquity"),
        "debt_to_equity": info.get("debtToEquity"),
        "revenue": info.get("totalRevenue"),
        "operating_margin": info.get("operatingMargins"),
        "profit_margin": info.get("profitMargins"),
        "dividend_yield": info.get("dividendYield"),
        "52week_high": info.get("fiftyTwoWeekHigh"),
        "52week_low": info.get("fiftyTwoWeekLow"),
        "50day_avg": info.get("fiftyDayAverage"),
        "200day_avg": info.get("twoHundredDayAverage"),
        "beta": info.get("beta"),
        "sector": info.get("sector"),
        "industry": info.get("industry"),
        "business_summary": info.get("longBusinessSummary", ""),
    }
    return data


def get_google_news(query, count=10):
    """Google News RSS로 뉴스 가져오기"""
    encoded_query = urllib.parse.quote(query + " 주가")
    url = f"https://news.google.com/rss/search?q={encoded_query}&hl=ko&gl=KR&ceid=KR:ko"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})

    try:
        raw = urllib.request.urlopen(req, timeout=10).read()
        root = ET.fromstring(raw)
        news_list = []
        for item in root.findall(".//item")[:count]:
            title = item.find("title").text or ""
            # 제목에서 " - 출처" 분리
            source = ""
            source_el = item.find("source")
            if source_el is not None:
                source = source_el.text or ""
            pub_date = ""
            pub_el = item.find("pubDate")
            if pub_el is not None:
                pub_date = pub_el.text or ""
            # 제목에서 출처 부분 제거 (예: "기사제목 - 한국경제" → "기사제목")
            clean_title = title.rsplit(" - ", 1)[0] if " - " in title else title
            news_list.append({
                "title": clean_title,
                "date": pub_date,
                "source": source,
            })
        return news_list
    except Exception:
        return []


def get_yfinance_news(ticker, count=10):
    """yfinance 뉴스 가져오기 (글로벌 종목)"""
    t = yf.Ticker(ticker)
    news_list = []
    for item in t.news[:count]:
        content = item.get("content", {})
        news_list.append({
            "title": content.get("title", ""),
            "date": content.get("pubDate", ""),
            "source": content.get("provider", {}).get("displayName", ""),
        })
    return news_list


def collect_all(query):
    """전체 데이터 수집 후 에이전트용 프롬프트 생성"""
    stock_data = get_stock_data(query)
    ticker = stock_data["ticker"]

    # 뉴스 수집 (Google News RSS)
    # 한국 종목은 한글 종목명으로, 글로벌 종목은 영문 이름으로 검색
    if ticker.endswith(".KS") or ticker.endswith(".KQ"):
        news = get_google_news(query)
    else:
        news = get_google_news(stock_data["name"])

    # 에이전트에 전달할 프롬프트 생성
    prompt = f"""## 분석 대상: {stock_data['name']} ({ticker})

### 기본 정보
- 업종: {stock_data.get('sector', 'N/A')} / {stock_data.get('industry', 'N/A')}
- 현재가: {stock_data.get('current_price', 'N/A')}
- 시가총액: {format_number(stock_data.get('market_cap'))}

### 밸류에이션
- PER (후행): {format_ratio(stock_data.get('per'))}
- PER (선행): {format_ratio(stock_data.get('forward_per'))}
- PBR: {format_ratio(stock_data.get('pbr'))}
- ROE: {format_pct(stock_data.get('roe'))}
- 부채비율: {format_ratio(stock_data.get('debt_to_equity'))}

### 수익성
- 매출: {format_number(stock_data.get('revenue'))}
- 영업이익률: {format_pct(stock_data.get('operating_margin'))}
- 순이익률: {format_pct(stock_data.get('profit_margin'))}
- 배당수익률: {format_pct(stock_data.get('dividend_yield'))}

### 주가 위치
- 52주 최고: {stock_data.get('52week_high', 'N/A')}
- 52주 최저: {stock_data.get('52week_low', 'N/A')}
- 50일 이동평균: {stock_data.get('50day_avg', 'N/A')}
- 200일 이동평균: {stock_data.get('200day_avg', 'N/A')}
- 베타: {stock_data.get('beta', 'N/A')}

### 기업 개요
{stock_data.get('business_summary', 'N/A')[:500]}

### 최근 뉴스
"""
    if news:
        for n in news:
            prompt += f"- [{n.get('source', '')}] {n.get('title', '')}\n"
    else:
        prompt += "- 뉴스 데이터 없음\n"

    prompt += "\n위 데이터를 바탕으로 당신의 전문 관점에서 이 종목을 분석하세요."

    return stock_data, prompt


def format_number(val):
    """큰 숫자를 읽기 쉽게 포맷"""
    if val is None:
        return "N/A"
    if val >= 1_000_000_000_000:
        return f"{val / 1_000_000_000_000:.1f}조"
    if val >= 100_000_000:
        return f"{val / 100_000_000:.0f}억"
    if val >= 1_000_000:
        return f"{val / 1_000_000:.1f}M"
    return str(val)


def format_ratio(val):
    if val is None:
        return "N/A"
    return f"{val:.2f}"


def format_pct(val):
    if val is None:
        return "N/A"
    return f"{val * 100:.1f}%"
