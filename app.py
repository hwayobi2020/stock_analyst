"""
멍거 격자모델 주식 분석 웹앱 - Flask 서버
"""
import asyncio
import os

from dotenv import load_dotenv
from flask import Flask, render_template, request, jsonify

from agents import run_all_agents, run_chief_analyst
from data_collector import collect_all

load_dotenv()

app = Flask(__name__)
API_KEY = os.getenv("ANTHROPIC_API_KEY")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    data = request.get_json()
    query = data.get("query", "").strip()

    if not query:
        return jsonify({"error": "종목명을 입력하세요."}), 400

    if not API_KEY:
        return jsonify({"error": "API 키가 설정되지 않았습니다."}), 500

    try:
        # 1. 데이터 수집
        stock_data, prompt = collect_all(query)

        # 2. 5개 에이전트 병렬 실행
        agent_results = asyncio.run(run_all_agents(API_KEY, prompt))

        # 3. 수석 분석관 취합
        final_report = run_chief_analyst(
            API_KEY, stock_data["name"], agent_results
        )

        return jsonify({
            "stock_name": stock_data["name"],
            "ticker": stock_data["ticker"],
            "agent_results": [
                {
                    "name": r["name"],
                    "icon": r["icon"],
                    "result": r["result"],
                }
                for r in agent_results
            ],
            "final_report": final_report,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=80, debug=False)
