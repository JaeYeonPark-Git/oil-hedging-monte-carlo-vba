# Oil Price Hedging Model with Barrier Options (VBA Monte Carlo Simulation)

![Excel VBA](https://img.shields.io/badge/Excel%20VBA-217346?style=flat&logo=microsoftexcel&logoColor=white)![Financial Engineering](https://img.shields.io/badge/Topic-Financial%20Engineering-blue)

**2023 Fall Semester**
Introduction to Financial Engineering (금융공학개론) Team Project

**Team:**
* Minje You, Sangik Lee, Taekjin Kim
* **Jae Yeon Park**

---

## 1. 🎯 프로젝트 개요 (Overview)

본 프로젝트는 유가 변동성(Volatility)에 따른 위험을 관리하면서 매월 안정적인 석유 공급을 확보하기 위해 설계된 금융 상품 모델입니다.

전쟁이나 지정학적 이슈로 인한 유가의 급격한 변동(Shock)으로부터 투자자를 보호하고, 안정적인 예산 편성을 돕는 것을 목표로 합니다.

## 2. 💡 모델 구조 (Model Structure)

이 모델은 **몬테카를로 시뮬레이션(Monte-Carlo Simulation)**을 기반으로 하며, 1년(12개월) 동안의 계약을 다룹니다.

### 주요 특징 (Features)
* **초기 3개월 (Initial Phase):**
    * 매월 고정 가격($72/gallon)으로 10갤런을 의무 구매하여 비용 안정성을 확보합니다.
* **이후 9개월 (Conditional Phase):**
    * **Case 1 (유가 < $90, Barrier 미만):** 추가 거래 없음.
    * **Case 2 (유가 >= $90, Barrier 초과):** 유가 상승에 유연하게 대응하기 위해 옵션 조건 발동 (Barrier Option 구조 활용).

### 파라미터 설정 (Settings)
* **Price Constraint:** Strike1보다 낮게 설정.
* **Spot vs Strike1:** Spot 가격은 Strike1보다 높아야 함.
* **Barrier:** 변동성이 높을수록 더 높은 Barrier와 Strike2를 설정하여 위험을 헷지함.

## 3. 💻 알고리즘 및 구현 (Algorithm)

본 프로젝트는 **Excel VBA**를 사용하여 **Geometric Brownian Motion (GBM)** 기반의 몬테카를로 시뮬레이션을 구현했습니다.

* **시뮬레이션 로직:**
    1.  `NPath`만큼의 시나리오 생성.
    2.  일별(Daily) 주가 이동: $S = S \times \exp((r - 0.5\sigma^2)dt + \sigma\sqrt{dt}Z)$
    3.  매월 만기 시점(`Maturity`)마다 Payoff 계산.
    4.  Barrier 도달 여부에 따른 Knock-out 또는 조건부 Payoff 계산.

## 4. 📊 시나리오 분석 결과 (Past Scenarios)

과거 데이터를 기반으로 모델의 수익성을 검증했습니다.

### 2022년 투자 시나리오
* **조건:** Volatility 25%, Barrier $94, Strike1 $70
* **결과:** Barrier가 4월 1일에 도달했음에도 불구하고, **$707.43의 순이익(Net Gain)** 발생.

### 2023년 투자 시나리오
* **조건:** Volatility 36%, Barrier $103, Strike1 $72
* **결과:** **$102.56의 순이익(Net Gain)** 발생.

> **결론:** 본 모델은 초기 비용 예측 가능성을 제공하고, 급격한 시장 변동 상황에서도 안정적인 수익 구조를 유지함을 확인했습니다.

## 5. 📁 리포지토리 구조 (Repository Structure)

```

/oil-hedging-monte-carlo-vba
├── /src
│   └── OilHedgingModel.bas       \# 몬테카를로 시뮬레이션 VBA 소스 코드
│
├── LICENSE                       \# MIT License
└── README.md                     \# 프로젝트 설명서

```

## 6. 🚀 실행 방법 (How to Run)

1.  Microsoft Excel을 실행하고 `Alt` + `F11`을 눌러 VBA 편집기를 엽니다.
2.  메뉴에서 `File` > `Import File`을 선택하고 `src/OilHedgingModel.bas` 파일을 불러옵니다.
3.  엑셀 시트에서 `=MonteCarloSimulation2(...)` 함수를 사용하여 파라미터를 입력하면 시뮬레이션이 실행됩니다.

---

## 📜 License

This project is licensed under the MIT License.
```
