# Oil Hedging Model (VBA Monte Carlo Simulation)

![VBA](https://img.shields.io/badge/VBA-Excel-217346?logo=microsoft-excel&logoColor=white)
![Financial Engineering](https://img.shields.io/badge/Topic-Financial%20Engineering-blue)

**2023 Fall Semester**
Introduction to Financial Engineering (금융공학개론) Team Project

**Team:**
* Minje You, Sangik Lee, Taekjin Kim
* **Jae Yeon Park**

---

## 1. 🎯 프로젝트 개요 (Overview)

[cite_start]**"Oil Hedging Model"**은 유가 변동성(Volatility)에 따른 위험을 관리하면서 매월 안정적인 석유 공급을 확보하기 위해 설계된 금융 상품 모델입니다[cite: 29, 46].

[cite_start]전쟁이나 지정학적 이슈로 인한 유가의 급격한 변동(Shock)으로부터 투자자를 보호하고, 안정적인 예산 편성을 돕는 것을 목표로 합니다 [cite: 20-25].

## 2. 💡 모델 구조 (Model Structure)

[cite_start]이 모델은 **몬테카를로 시뮬레이션(Monte-Carlo Simulation)**을 기반으로 하며, 1년(12개월) 동안의 계약을 다룹니다[cite: 62].

### 주요 특징 (Features)
* **초기 3개월 (Initial Phase):**
    * [cite_start]매월 고정 가격($72/gallon)으로 10갤런을 의무 구매하여 비용 안정성을 확보합니다[cite: 49].
* **이후 9개월 (Conditional Phase):**
    * **Case 1 (유가 < $90, Barrier 미만):** 추가 거래 없음.
    * [cite_start]**Case 2 (유가 >= $90, Barrier 초과):** 유가 상승에 유연하게 대응하기 위해 옵션 조건 발동 (Barrier Option 구조 활용) [cite: 50-54].

### 파라미터 설정 (Settings)
* **Price Constraint:** Strike1보다 낮게 설정.
* **Spot vs Strike1:** Spot 가격은 Strike1보다 높아야 함.
* [cite_start]**Barrier:** 변동성이 높을수록 더 높은 Barrier와 Strike2를 설정하여 위험을 헷지함 [cite: 33-41].

## 3. 💻 알고리즘 및 구현 (Algorithm)

본 프로젝트는 **Excel VBA**를 사용하여 **Geometric Brownian Motion (GBM)** 기반의 몬테카를로 시뮬레이션을 구현했습니다.

* **시뮬레이션 로직:**
    1.  `NPath`만큼의 시나리오 생성.
    2.  일별(Daily) 주가 이동: $S = S \times \exp((r - 0.5\sigma^2)dt + \sigma\sqrt{dt}Z)$
    3.  매월 만기 시점(`Maturity`)마다 Payoff 계산.
    4.  [cite_start]Barrier 도달 여부에 따른 Knock-out 또는 조건부 Payoff 계산 [cite: 72-137].

## 4. 📊 시나리오 분석 결과 (Past Scenarios)

과거 데이터를 기반으로 모델의 수익성을 검증했습니다.

### 2022년 투자 시나리오
* **조건:** Volatility 25%, Barrier $94, Strike1 $70
* [cite_start]**결과:** Barrier가 4월 1일에 도달했음에도 불구하고, **$707.43의 순이익(Net Gain)** 발생 [cite: 143-151].

### 2023년 투자 시나리오
* **조건:** Volatility 36%, Barrier $103, Strike1 $72
* [cite_start]**결과:** **$102.56의 순이익(Net Gain)** 발생 [cite: 154-162].

> **결론:** 본 모델은 초기 비용 예측 가능성을 제공하고, 급격한 시장 변동 상황에서도 안정적인 수익 구조를 유지함을 확인했습니다.

## 5. 📁 리포지토리 구조 (Repository Structure)
