# Oil Price Hedging Model with Barrier Options (VBA Monte Carlo Simulation)

![Excel VBA](https://img.shields.io/badge/Excel%20VBA-217346?style=flat&logo=microsoftexcel&logoColor=white&cache=123)
![Financial Engineering](https://img.shields.io/badge/Topic-Financial%20Engineering-blue)

**Course Project: Introduction to Financial Engineering (UNIST, Fall 2023)**

**Team:**
* Minje You, Sangik Lee, Taekjin Kim
* **Jae Yeon Park**

---

## 🎯 프로젝트 개요 (Overview)

본 프로젝트는 유가 변동성(Volatility)에 따른 위험을 관리하면서 매월 안정적인 석유 공급을 확보하기 위해 설계된 금융 상품 모델입니다.

전쟁이나 지정학적 이슈로 인한 유가의 급격한 변동(Shock)으로부터 투자자를 보호하고, 안정적인 예산 편성을 돕는 것을 목표로 합니다. 특히, 몬테카를로 시뮬레이션을 통해 다양한 시장 시나리오를 예측하고, Barrier Option 구조를 활용하여 유연한 리스크 헷징 전략을 제시합니다.

## 💡 모델 구조 (Model Structure)

이 모델은 **몬테카를로 시뮬레이션(Monte-Carlo Simulation)**을 기반으로 하며, 1년(12개월) 동안의 계약을 다룹니다.

### 주요 특징 (Features)
* **초기 3개월 (Initial Phase - Fixed Rate):**
    * 매월 고정 가격($72/gallon)으로 10갤런을 의무 구매하여 비용 안정성을 확보합니다.
    * 초기 시장 진입 시점의 불확실성을 최소화하고 안정적인 물량 확보를 우선시합니다.
* **이후 9개월 (Conditional Phase - Barrier Option):**
    * 시장 상황에 따라 유연하게 대응하는 조건부 옵션 구조를 가집니다.
    * **Case 1 (유가 < Barrier):** 배리어(Barrier) 가격 미만일 경우, 옵션이 유지되며 만기 시점의 주가에 따라 Payoff가 결정됩니다.
    * **Case 2 (유가 >= Barrier):** 배리어(Barrier) 가격 이상으로 상승할 경우(Knock-out), 추가 거래를 중단하여 손실을 방지하거나 리스크를 관리합니다.

### 파라미터 설정 (Settings)
* **Price Constraint:** Strike1보다 낮게 설정하여 구매자에게 유리한 가격 조건을 탐색합니다.
* **Spot vs Strike1:** Spot 가격은 Strike1보다 높아야 옵션 행사 가치가 발생하도록 설정합니다.
* **Barrier:** 변동성이 높을수록 더 높은 Barrier와 Strike2를 설정하여 위험을 헷지합니다. 높은 변동성은 Barrier 도달 확률을 높이므로, 이에 대한 보상으로 Strike 가격을 조정합니다.

## 💻 알고리즘 및 구현 (Algorithm)

본 프로젝트는 **Excel VBA**를 사용하여 **Geometric Brownian Motion (GBM)** 기반의 몬테카를로 시뮬레이션을 구현했습니다.

* **핵심 로직 (Simulation Logic):**
    1.  **시나리오 생성:** `NPath`만큼의 주가 경로(Path)를 생성하여 다양한 시장 상황을 시뮬레이션합니다.
    2.  **주가 이동 (GBM):** 매일(Daily) 주가는 기하학적 브라운 운동을 따르며 변동합니다.
        * 수식: $S_{t+dt} = S_t \times \exp((r - 0.5\sigma^2)dt + \sigma\sqrt{dt}Z)$
        * 여기서 $Z$는 표준 정규 분포를 따르는 확률 변수입니다.
    3.  **월별 Payoff 계산:**
        * **Month 1-3:** 고정된 Strike1 가격으로 매월 10단위 수익 계산.
        * **Month 4-12:**
            * 매일 주가를 모니터링하여 Barrier 도달 여부 확인 (Knock-out).
            * 만기 시점(`Maturity`)에 주가가 Strike1보다 높으면 Strike1 가격으로 10단위 수익 실현.
            * 주가가 Strike1보다 낮으면 Strike2 가격으로 20단위 수익 실현 (저가 매수 기회 확대).
    4.  **할인 (Discounting):** 각 시점의 Payoff를 현재 가치(Present Value)로 할인하여 합산합니다.
    5.  **평균값 도출:** 모든 시나리오의 Payoff 평균을 계산하여 옵션의 공정 가치(Fair Value)를 도출합니다.

## 📊 시나리오 분석 결과 (Past Scenarios)

과거 데이터를 기반으로 모델의 수익성을 검증했습니다.

### 2022년 투자 시나리오
* **조건:** Volatility 25%, Barrier $94, Strike1 $70
* **결과:** Barrier가 4월 1일에 도달했음에도 불구하고, **$707.43의 순이익(Net Gain)** 발생.
* **분석:** 높은 변동성에도 불구하고 초기 고정 수익과 배리어 도달 전까지의 수익이 전체 포트폴리오를 방어했습니다.

### 2023년 투자 시나리오
* **조건:** Volatility 36%, Barrier $103, Strike1 $72
* **결과:** **$102.56의 순이익(Net Gain)** 발생.
* **분석:** 변동성이 더욱 커진 상황에서도 안정적인 수익 구조를 유지하며, 모델의 리스크 헷징 성능을 입증했습니다.

> **결론:** 본 모델은 초기 비용 예측 가능성을 제공하고, 급격한 시장 변동 상황에서도 안정적인 수익 구조를 유지함을 확인했습니다.

## 📁 리포지토리 구조 (Repository Structure) 

```

/oil-hedging-monte-carlo-vba
├── OilHedgingModel.bas       \# 몬테카를로 시뮬레이션 VBA 소스 코드
├── LICENSE                       \# MIT License
└── README.md                     \# 프로젝트 설명서

```
## 🚀 실행 방법 (How to Run)

1.  Microsoft Excel을 실행하고 `Alt` + `F11`을 눌러 VBA 편집기를 엽니다.
2.  메뉴에서 `File` > `Import File`을 선택하고 `OilHedgingModel.bas` 파일을 불러옵니다. (또는 모듈을 추가하여 아래 코드를 붙여넣으세요.)
3.  엑셀 시트의 임의의 셀에서 `=MonteCarloSimulation2(...)` 함수를 입력합니다.
    * **입력 파라미터 예시:** Spot, Strike1, Strike2, Barrier, Volatility, IR, Maturity1~12 (날짜), NPath, AsofDate
4.  결과값이 계산되어 셀에 표시됩니다. (시뮬레이션 횟수 `NPath`가 클수록 계산 시간이 길어질 수 있습니다.)

---

## 📜 License

This project is licensed under the MIT License.
```
