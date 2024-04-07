what does this cshar^function ? protected void postLastImport_Click(object Sender, ButtonEventArgs e)
    {
      string path1 = HttpContext.Current.Server.MapPath("uploadfiles");
      if (!Directory.Exists(path1))
        Directory.CreateDirectory(path1);
      if (!this.Upload1.HasFiles)
        return;
      for (int index1 = 0; index1 < this.Upload1.PostedFiles.Count; ++index1)
      {
        string str1 = this.Upload1.PostedFiles[index1].FileName;
        int num1 = str1.LastIndexOf("\\");
        if (num1 > -1)
          str1 = str1.Substring(num1 + 1);
        this.Upload1.PostedFiles[index1].SaveAs(path1 + "\\" + str1);
        if (str1.Length > 0)
        {
          string[] files = Directory.GetFiles(this.Request.PhysicalApplicationPath, "*.xls");
          Directory.GetFiles(this.Request.PhysicalApplicationPath, "*.xlsx");
          foreach (string path2 in files)
            File.Delete(path2);
          string str2 = path1 + "\\" + str1;
          MySqlTransaction tr = this.sqlquery.OpenTransaction();
          try
          {
            this.Session.Add("ERRORIMPORT", (object) "");
            this.Session["ERRORIMPORT"] = (object) "";
            string text1 = this.tx_titolo2.Text;
            string text2 = this.tx_sheat2.Text;
            string str3 = this.ValidateImport(text1, text2);
            if (str3 == "OK")
            {
              OleDbConnection selectConnection = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" + str2 + ";Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"");
              OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter("select * from ['" + text2 + "$']", selectConnection);
              System.Data.DataTable dataTable = new System.Data.DataTable();
              oleDbDataAdapter.Fill(dataTable);
              string str4 = "";
              string str5 = "";
              string str6 = "";
              string str7 = "0";
              string str8 = "";
              string str9 = "";
              string str10 = "";
              string str11 = "";
              string str12 = "0";
              string str13 = "";
              string str14 = "";
              string str15 = "";
              string str16 = "";
              string str17 = "";
              string str18 = "";
              string str19 = "";
              string str20 = "";
              DateTime dateTime1 = new DateTime();
              DateTime dateTime2 = new DateTime();
              bool flag1 = false;
              bool flag2 = false;
              bool flag3 = false;
              bool flag4 = false;
              bool flag5 = false;
              bool flag6 = false;
              bool flag7 = false;
              bool flag8 = false;
              bool flag9 = false;
              bool flag10 = false;
              string str21 = "";
              bool flag11 = false;
              bool flag12 = false;
              string str22 = "";
              bool flag13 = false;
              bool flag14 = false;
              string str23 = "";
              bool flag15 = false;
              bool flag16 = false;
              string str24 = "";
              bool flag17 = false;
              bool flag18 = false;
              string str25 = "";
              bool flag19 = false;
              bool flag20 = false;
              string str26 = "";
              bool flag21 = false;
              bool flag22 = false;
              bool flag23 = false;
              bool flag24 = false;
              string str27 = "";
              string str28 = "";
              DateTime dateTime3 = new DateTime();
              int num2 = 0;
              int num3 = 0;
              DateTime dateTime4 = new DateTime();
              DateTime dateTime5 = new DateTime();
              dateTime4 = DateTime.Now;
              dateTime5 = Convert.ToDateTime("1001-01-01 00:00:00");
              string sql1 = "select ifnull(max(idimport),0)+1 from c_importazioni";
              DataSet dataSet1 = this.sqlquery.creaDatasetTransaction(tr, sql1);
              if (dataSet1.Tables[0].Rows.Count > 0)
                num2 = Convert.ToInt32(dataSet1.Tables[0].Rows[0][0]);
              for (int index2 = 0; index2 < dataTable.Rows.Count; ++index2)
              {
                if (index2 > 0)
                {
                  ++num3;
                  if (!DBNull.Value.Equals(dataTable.Rows[index2][10]))
                    str13 = dataTable.Rows[index2][10].ToString().Replace("'", "''");
                  if (!DBNull.Value.Equals(dataTable.Rows[index2][10]))
                  {
                    string sql2 = "update c_programmatori set onlineagg = 1,idaggiornamento = 1, aggiornato = 1 where idimport = '" + num2.ToString() + "' and serial ='" + str13 + "' ";
                    this.sqlquery.salvaDataset(tr, sql2);
                    string str29 = "insert into c_aggiornamento values(@idaggiornamento , @idimport, @idriga, @locationkey,@location,@customer,@reparto,@area,@regione,@indirizzo,@citta,@provincia, @zipcode, @country,@maps,@cshregione,@cshterritory,@modelnumber,@modeldesc,@serial, @totalqty,@lastdelivery,@materialkey,@salesrepname,@versione,@codicerep,@aggiornatoda,@data_aggiornamento)";
                    str29 = "insert into c_aggiornamento VALUES ( 1, '" + num2.ToString() + "' , '" + num3.ToString() + "' , '" + str4 + "' , '" + str5 + "' , '" + str6 + "' , '" + str18 + "' , '" + str10 + "' , '" + str9 + "' , '" + str16 + "' , '" + str7 + "' , '' , '' , '' , '' , '' , '" + str11 + "' , '" + str12 + "' , '" + str13 + "' , '" + str14 + "' , '' , '' , '' , '' , ''  )";
                  }
                  if (!DBNull.Value.Equals(dataTable.Rows[index2][11]))
                  {
                    string sql3 = "update c_programmatori set onlinecheck = 1, checkannuale = 1 where idimport = '" + num2.ToString() + "' and serial ='" + str13 + "' ";
                    this.sqlquery.salvaDataset(tr, sql3);
                  }
                  int num4 = 0;
                  string sql4 = "select * from c_prodottiprogrammatori where modeldesc = '" + str12 + "'";
                  DataSet dataSet2 = this.sqlquery.creaDatasetTransaction(tr, sql4);
                  if (dataSet2.Tables[0].Rows.Count > 0)
                    num4 = Convert.ToInt32(dataSet2.Tables[0].Rows[0]["idprodotto"]);
                  string[] strArray = new string[109];
                  strArray[0] = "INSERT INTO c_datilastprog VALUES ('";
                  strArray[1] = num2.ToString();
                  strArray[2] = "',  '";
                  strArray[3] = num3.ToString();
                  strArray[4] = "', '";
                  strArray[5] = str4;
                  strArray[6] = "', '";
                  strArray[7] = str5;
                  strArray[8] = "', '";
                  strArray[9] = str6;
                  strArray[10] = "', '";
                  strArray[11] = str7;
                  strArray[12] = "', '";
                  strArray[13] = str8;
                  strArray[14] = "', '";
                  strArray[15] = str9;
                  strArray[16] = "', '";
                  strArray[17] = str10;
                  strArray[18] = "', '";
                  strArray[19] = str11;
                  strArray[20] = "', '";
                  strArray[21] = str12;
                  strArray[22] = "', '";
                  strArray[23] = str13;
                  strArray[24] = "', '";
                  strArray[25] = str14;
                  strArray[26] = "', '";
                  strArray[27] = str15;
                  strArray[28] = "', '";
                  strArray[29] = str16;
                  strArray[30] = "', '";
                  strArray[31] = str17;
                  strArray[32] = "', '";
                  strArray[33] = str18;
                  strArray[34] = "', '";
                  strArray[35] = str19;
                  strArray[36] = "', '";
                  strArray[37] = str20;
                  strArray[38] = "', '";
                  strArray[39] = dateTime1.ToString(ConfigurationSettings.AppSettings["datetimeFormatSt"].ToString());
                  strArray[40] = "', '";
                  strArray[41] = dateTime2.ToString(ConfigurationSettings.AppSettings["datetimeFormatSt"].ToString());
                  strArray[42] = "', '";
                  int int32 = Convert.ToInt32(flag1);
                  strArray[43] = int32.ToString();
                  strArray[44] = "', '";
                  int32 = Convert.ToInt32(flag2);
                  strArray[45] = int32.ToString();
                  strArray[46] = "', '";
                  int32 = Convert.ToInt32(flag3);
                  strArray[47] = int32.ToString();
                  strArray[48] = "', '";
                  int32 = Convert.ToInt32(flag4);
                  strArray[49] = int32.ToString();
                  strArray[50] = "', '";
                  int32 = Convert.ToInt32(flag5);
                  strArray[51] = int32.ToString();
                  strArray[52] = "', '";
                  int32 = Convert.ToInt32(flag6);
                  strArray[53] = int32.ToString();
                  strArray[54] = "', '";
                  int32 = Convert.ToInt32(flag7);
                  strArray[55] = int32.ToString();
                  strArray[56] = "', '";
                  int32 = Convert.ToInt32(flag8);
                  strArray[57] = int32.ToString();
                  strArray[58] = "', '";
                  int32 = Convert.ToInt32(flag9);
                  strArray[59] = int32.ToString();
                  strArray[60] = "', '";
                  int32 = Convert.ToInt32(flag10);
                  strArray[61] = int32.ToString();
                  strArray[62] = "', '";
                  strArray[63] = str21;
                  strArray[64] = "', '";
                  int32 = Convert.ToInt32(flag11);
                  strArray[65] = int32.ToString();
                  strArray[66] = "', '";
                  int32 = Convert.ToInt32(flag12);
                  strArray[67] = int32.ToString();
                  strArray[68] = "', '";
                  strArray[69] = str22;
                  strArray[70] = "', '";
                  int32 = Convert.ToInt32(flag13);
                  strArray[71] = int32.ToString();
                  strArray[72] = "', '";
                  int32 = Convert.ToInt32(flag14);
                  strArray[73] = int32.ToString();
                  strArray[74] = "', '";
                  strArray[75] = str23;
                  strArray[76] = "', '";
                  int32 = Convert.ToInt32(flag15);
                  strArray[77] = int32.ToString();
                  strArray[78] = "', '";
                  int32 = Convert.ToInt32(flag16);
                  strArray[79] = int32.ToString();
                  strArray[80] = "', '";
                  strArray[81] = str24;
                  strArray[82] = "', '";
                  int32 = Convert.ToInt32(flag17);
                  strArray[83] = int32.ToString();
                  strArray[84] = "', '";
                  int32 = Convert.ToInt32(flag18);
                  strArray[85] = int32.ToString();
                  strArray[86] = "', '";
                  strArray[87] = str25;
                  strArray[88] = "', '";
                  int32 = Convert.ToInt32(flag19);
                  strArray[89] = int32.ToString();
                  strArray[90] = "', '";
                  int32 = Convert.ToInt32(flag20);
                  strArray[91] = int32.ToString();
                  strArray[92] = "', '";
                  strArray[93] = str26;
                  strArray[94] = "', '";
                  int32 = Convert.ToInt32(flag21);
                  strArray[95] = int32.ToString();
                  strArray[96] = "', '";
                  int32 = Convert.ToInt32(flag22);
                  strArray[97] = int32.ToString();
                  strArray[98] = "', '";
                  int32 = Convert.ToInt32(flag23);
                  strArray[99] = int32.ToString();
                  strArray[100] = "', '";
                  int32 = Convert.ToInt32(flag24);
                  strArray[101] = int32.ToString();
                  strArray[102] = "', '";
                  strArray[103] = str27;
                  strArray[104] = "', '";
                  strArray[105] = str28;
                  strArray[106] = "', '";
                  strArray[107] = dateTime3.ToString(ConfigurationSettings.AppSettings["datetimeFormatSt"].ToString());
                  strArray[108] = "' )  ";
                  string sql5 = string.Concat(strArray);
                  this.sqlquery.salvaDataset(tr, sql5);
                }
                else
                {
                  DateTime dateTime6 = new DateTime();
                  DateTime dateTime7 = new DateTime();
                  DateTime dateTime8 = new DateTime();
                  DateTime dateTime9 = new DateTime();
                }
              }
              this.Session["ERRORIMPORT"] = (object) ("Importazione Eseguita\r\nImportatate " + (dataTable.Rows.Count - 1).ToString() + " righe");
              this.lbl_errore2.Text = this.Session["ERRORIMPORT"].ToString();
            }
            else
              this.lbl_errore2.Text = str3;
          }
          catch (Exception ex)
          {
            ((DbTransaction) tr).Rollback();
            this.lbl_errore2.Text = ex.Message + ex.StackTrace;
          }
          finally
          {
            this.sqlquery.CloseTransaction(tr);
            ((DbTransaction) tr).Dispose();
          }
        }
      }
