using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.Utils.Extensions;

namespace ExcelReaderWriter
{
    internal static class MetodoExtensao
    {
        public static string RemoverAcentuacao(this string text)
        {
            return new string(text
                .Normalize(NormalizationForm.FormD)
                .Where(ch => char.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                .ToArray());
        }

        public static string RemoverPreposicao(this string text)
        {
            return text.ToLower().Replace(" da ", " ")
                .Replace(" de ", " ")
                .Replace(" di ", " ")
                .Replace(" do ", " ")
                .Replace(" du ", " ")
                .Replace(" d'", " ");
        }

        public static string RemoverEstado(this string text)
        {
            List<string> estadosSigla = new List<string>()
            {
                "ac","al", "ap", "am", "ba", "ce", "df", "es", "go", "ma"
                , "mt", "ms", "mg", "pa", "pb", "pr", "pe", "pi", "rj", "rn"
                ,"rs", "ro", "rr", "sc", "sp", "se", "to"
            };
            var teste = text.Split('-');
            var result = new StringBuilder();

            teste.ForEach(x =>
            {
                if (!estadosSigla.Contains(x.ToLower().Replace(" ", "")))
                {
                    result.Append(x);
                }
            });

            return result.ToString();
        }
    }
}