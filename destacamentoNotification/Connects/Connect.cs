namespace notificacaoSemanalTestes.Connects
{
    class Connect
    {
        public static Connect_HT_server HTlocalConnect = new Connect_HT_server(Security.settings.ht_HName, Security.settings.ht_DBName, Security.settings.ht_UName, Security.settings.ht_Pass);
        public static Connect_HT_server SVlocalConnect = new Connect_HT_server(Security.settings.ht_HName, "secretariaVirtual", Security.settings.ht_UName, Security.settings.ht_Pass);

        public static void closeAll()
        {
            HTlocalConnect.Conn.Close();
            SVlocalConnect.Conn.Close();
        }
    }
}