using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;

namespace NBABet
{
    public static class BancoDados
    {
        public static string ConsultaString(string instrucao, Dictionary<string, object> parametros = null)
        {
            using (var connection = new SqliteConnection("Data Source=Banco_Dados\\NBAStats.db"))
            {
                connection.Open();

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = instrucao;

                    if (parametros != null)
                    {
                        foreach (var param in parametros)
                        {
                            command.Parameters.AddWithValue(param.Key, param.Value);
                        }
                    }

                    return command.ExecuteScalar()?.ToString();
                }
            }
        }

        public static int? ConsultaInt(string instrucao, Dictionary<string, object> parametros = null)
        {
            using (var connection = new SqliteConnection("Data Source=Banco_Dados\\NBAStats.db"))
            {
                connection.Open();

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = instrucao;

                    if (parametros != null)
                    {
                        foreach (var param in parametros)
                        {
                            command.Parameters.AddWithValue(param.Key, param.Value);
                        }
                    }

                    return (int?)command.ExecuteScalar();
                }
            }
        }

        public static SqliteDataReader ConsultaTabela(string instrucao, Dictionary<string, object> parametros = null)
        {
            using (var connection = new SqliteConnection("Data Source=Banco_Dados\\NBAStats.db"))
            {
                connection.Open();

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = instrucao;

                    if (parametros != null)
                    {
                        foreach (var param in parametros)
                        {
                            command.Parameters.AddWithValue(param.Key, param.Value);
                        }
                    }

                    return command.ExecuteReader();
                }
            }
        }

        public static int Insere(string instrucao, Dictionary<string, object> parametros = null)
        {
            try
            {
                using (var connection = new SqliteConnection("Data Source=Banco_Dados\\NBAStats.db"))
                {
                    connection.Open();

                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = instrucao;

                        if (parametros != null)
                        {
                            foreach (var param in parametros)
                            {
                                command.Parameters.AddWithValue(param.Key, param.Value);
                            }
                        }

                        return command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
