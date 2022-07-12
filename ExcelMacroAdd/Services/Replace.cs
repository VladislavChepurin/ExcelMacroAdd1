namespace ExcelMacroAdd.Servises
{
    internal static class Replace
    {
        /// <summary>
        /// Функция замены для вставки вендора и запроса из XML
        /// </summary>
        /// <param name="mReplase"></param>
        /// <returns></returns>
        public static string RepleceVendorTable(string mReplase)                          // Функция замены // индус заплачит от умиления IEK ВА47 - кирилица, IEK BA47М - латиница Переписать!!!
        {
            //Функция несет двойную задачу, это неправильно.
            return mReplase.Replace("IEK ВА47", "IEK").Replace("IEK BA47М", "IEK").Replace("EKF PROxima", "EKF").Replace("EKF AVERS", "EKF").Replace("Schneider", "SE").Replace("IEK Armat", "IEK");
        }
        /// <summary>
        /// Функция замены для запроса SQL
        /// </summary>
        /// <param name="mReplase"></param>
        /// <returns></returns>
        public static string FuncReplece(string mReplase)                                 // Функция замены // индус заплачит от умиления IEK ВА47 - кирилица, IEK BA47М - латиница Переписать!!!
        {
            return mReplase.Replace("IEK ВА47", "iek_va47").Replace("IEK BA47М", "iek_va47m").Replace("EKF PROxima", "ekf_proxima").Replace("ABB", "abb").Replace("EKF AVERS", "ekf_avers").
                Replace("KEAZ", "keaz").Replace("DKC", "dkc").Replace("DEKraft", "dekraft").Replace("Schneider", "schneider").Replace("TDM", "tdm").Replace("IEK Armat", "iek_armat");
        }
        /// <summary>
        /// Фнкцция замены для ВПР при считывании
        /// </summary>
        /// <param name="mReplase"></param>
        /// <param name="rows"></param>
        /// <returns></returns>
        public static string VprFormulaReplace(string mReplase, int rows)
        {
            return mReplase.Replace("=ВПР(A"+ rows.ToString(), "=ВПР(A{0}");
        }
    }
}
