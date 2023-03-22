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
- [X] Take random positions at random moments and update balances

### Milestones :

- [X] Data tab 👉 Dynamic candlesticks chart (~ 20 candles) 1 min + Same but linear and 1 sec.
- [X] Balances tab 👉 Input field to enter Binance testnet API keys + display balances + button to refresh / auto-refresh
- [X] Trading tab 👉 Buttons to buy / sell / TP / SL (if possible) + Auto strategies : Mean-reversion & Momentum
- [X] Prediction tab  👉 KNN 
- [X] About tab 👉 Link to this repo, details about the project, share buttons etc.

## Demo API Keys

Here are some API keys created with disposable GitHub accounts. These keys are available to everyone, to avoid creating an account on Binance Testnet Vision (https://testnet.binance.vision/). Please take care of them and don't abuse the initial balance.

| Account       | Keys                                                             |
|---------------|------------------------------------------------------------------|
| #1 API Key    | c4sw4elBs01FWXlNt1gaZiq4tORKCag6PS4SLTya1ygObExJMV2uq1F0lhJ1G2Oc |
| #1 Secret Key | Ct3wt0TK5upZB2plITMtOStLiwqlo5qKnVZZttcy3aeggb353y52uPQt1OX3Fjge |
| #2 API Key    | bkNMUJq3pAcjl7oDQYvpSogYVAp2l31Q88MJK8P7iyQv5Z3rEGMAy6pkOoYJzvkO |
| #2 Secret Key | 9rgEmbaIhk7TMTiZoVg8kOWit1xJyP4s6QOenVFaJZ8gxElQXPoeDLNYI7knSqFf |
| #3 API Key    | uc4sFtIm5CXTMyrexAp4JV7rcthCULrAQhkzttKHPNUksAgx6iOOWEIphSxbGPWp |
| #3 Secret Key | OjWejYkPkmT8K3ZjiPjjA2rCpYXoXRPf6jkGP6eBjYIO11vX5riJZaWySlTgYmvT |
| #4 API Key    | JTPIHJrQI8PL79Wmmq1Z3fDVGFrkkuK3tJICRAmEdYGYuy7fFPV0MKVMbI1eFaOV |
| #4 Secret Key | DiSZkeAJZCoVT0OJEWoQrp0m16ztpfZ8QMxt90HkZk9HIrn93tX0u7E1yo2wqToo |
| #5 API Key    | JEaBeUXBMjE2TYCzuEvTIs1xF6okEB9e3zUf8wJgUlgYcn4Da7NXHnvU6PHZ3VxO |
| #5 Secret Key | kPhBwuFvNDx7SDACExVde14qk4Cxc3NpgGphJ4c0JDGjDs0uD3qZRGtbc0578L8G |

