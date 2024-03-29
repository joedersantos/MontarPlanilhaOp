﻿using System;
using System.Globalization;

namespace MontarPlanilha.Utils
{
    public static class StringExtension
    {
        private static readonly CultureInfo CultureInfo = CultureInfo.CreateSpecificCulture("en-US");
        public static int ToInt(this string value, int linha)
        {
            if (int.TryParse(value.Trim(), out int idOut))
                return idOut;
            else
                throw new Exception($"Erro na conversão de dados comuna {linha}. Verifique o formato do arquivo .csv !");
        }

        public static decimal ToDecimal(this string value, int linha)
        {
            value = value.Trim().Replace(" ", "");
            value = value.Replace(",", ".");
            if (!decimal.TryParse(value, out decimal outDecimal))               
                throw new Exception($"Erro na conversão de dados comuna {linha}. Verifique o formato do arquivo .csv !");

            return decimal.Parse(value, CultureInfo);
            //return decimal.Parse(value, CultureInfo.InvariantCulture);
        }
    }
}
