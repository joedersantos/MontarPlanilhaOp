using MontarPlanilha.Utils;
using System;

namespace MontarPlanilha
{
    public class Operacao
    {
        public Operacao(string ordem)
        {
            var colunas = ordem.Split(";");
            AddDate(colunas[0]);
            Id = colunas[1].ToInt(1);
            Symbol = colunas[2].Trim();
            Type = colunas[3].Trim();
            Direction = colunas[4].Trim();
            Volume = colunas[5].ToInt(5);
            Price = colunas[6].ToDecimal(6);
            AddProfit(colunas[10].Trim());
        }

        #region privados Add
        private void AddDate(string date)
        {
            var axDt = date.Split(".");
            var axDayTime = axDt[2].Split(" ");
            var axTime = axDayTime[1].Trim().Split(":");

            Date = new DateTime(axDt[0].ToInt(0), axDt[1].ToInt(0), axDayTime[0].ToInt(0), axTime[0].ToInt(0), axTime[1].ToInt(0), axTime[2].ToInt(0));
        }
        private void AddProfit(string profit)
        {
            var ax = profit.ToDecimal(10);
            Profit = ax / 1000;
        }
        #endregion
        public int Id { get; private set; }
        public DateTime Date { get; private set; }
        public string Symbol { get; private set; }
        /// <summary>
        /// Tipo da operação
        /// </summary>
        /// <remarks>Buy or Sell</remarks>
        public string Type { get; private set; }
        /// <summary>
        /// Ordem de entrada ou saida
        /// </summary>
        /// <remarks>In or Out</remarks>
        public string Direction { get; private set; }

        /// <summary>
        /// Quatidade de lotes
        /// </summary>
        public int Volume { get; private set; }
        public decimal Price { get; private set; }
        /// <summary>
        /// resusltado da operação
        /// vamos origenal do MT5 dividido por 1000
        /// </summary>
        public decimal Profit { get; private set; }


        public bool CheckOut(Operacao opOut)
        {
            decimal axProft;
            if (Type == "buy")
                axProft = opOut.Price - Price;
            else
                axProft = Price - opOut.Price;

            return axProft == opOut.Profit;
        }

    }
}
