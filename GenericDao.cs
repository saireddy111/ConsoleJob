using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleJob
{
    class GenericDao
    {
        public string Invoke(string procedureName, Hashtable hash)
        {
            IDictionaryEnumerator en = hash.GetEnumerator();
            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DataContext"].ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(procedureName, conn))
                {
                    try
                    {
                        cmd.CommandType = CommandType.StoredProcedure;                       
                        cmd.Connection = conn;
                        cmd.CommandTimeout = 0;
                        while (en.MoveNext())
                        {
                            if (en.Key.ToString() == "@p_out_response_data")
                            {
                                SqlParameter p_out_response_data = new SqlParameter("@p_out_response_data", SqlDbType.NVarChar, -1)
                                {
                                    Direction = ParameterDirection.Output
                                };
                                cmd.Parameters.Add(p_out_response_data);
                            }
                            else
                            {
                                cmd.Parameters.AddWithValue(en.Key.ToString(), en.Value);
                            }
                        }

                        conn.Open();
                        
                        cmd.ExecuteNonQuery();
                        return cmd.Parameters["@p_out_response_data"].Value.ToString();
                        
                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.Message);
                        return "{\"STATUS\":\"ERROR\",\"REMARKS\":\"" + ex.Message + "\"}";
                    }
                    finally
                    {
                        conn.Close();
                        conn.Dispose();
                    }
                }
            }
        }
    }
}
