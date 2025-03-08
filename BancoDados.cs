using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;

namespace NBABet
{
    public static class BancoDados
    {
        static string connectionString = "C:\\Users\\ricardo.queiroz\\source\\repos\\NBABet\\Banco_Dados\\NBAStats.db";

        public static string ConsultaString(string instrucao, Dictionary<string, object> parametros = null)
        {
            using (var connection = new SqliteConnection($"Data Source={connectionString}"))
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

        public static long? ConsultaInt(string instrucao, Dictionary<string, object> parametros = null)
        {
            using (var connection = new SqliteConnection($"Data Source={connectionString}"))
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

                    return (long?)command.ExecuteScalar();
                }
            }
        }

        public static List<T> ConsultaTabela<T>(string instrucao, Func<SqliteDataReader, T> map, Dictionary<string, object> parametros = null)
        {
            List<T> resultados = new List<T>();

            using (var connection = new SqliteConnection($"Data Source={connectionString}"))
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

                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            resultados.Add(map(reader));
                        }
                    }
                }
            }

            return resultados;
        }

        public static int Insere(string instrucao, Dictionary<string, object> parametros = null)
        {
            try
            {
                using (var connection = new SqliteConnection($"Data Source={connectionString}"))
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
