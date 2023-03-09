![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
[![Generic badge](https://img.shields.io/badge/VBA-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white-<COLOR>.svg)](https://shields.io/)

![](https://i.ibb.co/3kvvLLJ/Excelliarmus.jpg)

## What is Excelliarmus ?
Excelliarmus is an open-source trading software, currently under development, that will allow semi-automated trading of financial assets, first crypto-currencies, as the data is easier to access, then other assets maybe. The technical constraint is that the project will be coded in VBA.

## What is the goal of this repo?
The goal of this repo is to provide a monitoring of the project progress, by setting small goals, and trying to implement the most features in a given time.

---

### List of small features to implement sorted by priority:

- [X] Create basic UserForm / mock programs
- [X] Button to show current BTC's price (HTTP request to Binance's API)
- [X] Download historical data (https://api.binance.com/api/v3/klines?symbol=BTCUSDT&interval=1h)
- [X] Display financial data as a chart
- [X] Make repeted requests and make chart update dynamically
- [X] Connect to Binance's subnet (Paper)
- [X] Display Binance's Paper accounts
- [ ] Take random positions at random moments and update balances

### Milestones :

- [X] Data tab 👉 Dynamic candlesticks chart (~ 20 candles) 1 min + Same but linear and 1 sec.
- [/] Balances tab 👉 Input field to enter Binance testnet API keys + display balances + button to refresh / auto-refresh
- [ ] Trading tab 👉 Buttons to buy / sell / TP / SL (if possible) + Auto strategies : Mean-reversion & Momentum
- [ ] Prediction tab (if enough time) 👉 SVM from scratch (functions .fit and .predict) or from Python Server (HTTP req. to API) + (if time) connection between predictions and actions on exchange(s)
- [X] About tab 👉 Link to this repo, details about the project, share buttons etc.
