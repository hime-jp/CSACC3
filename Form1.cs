using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Xml.Serialization;

namespace CSACC3
{
    public partial class Form1 : Form
    {

        string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=CSDB3.accdb";
        string strAccessSelect = "SELECT * FROM waza";
        string strAccessWhere;
        string strAccessInsert;
        string strAccessDelete;
        string strAccessUpdate;

        OleDbConnection myAccessConn = null;

        public Form1()
        {
            InitializeComponent();
        }

        //検索ボタンの処理

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed connection. " + ex.Message;
                return;
            }

            try
            {
                if (comboBox1.Text.Length != 0)
                {
                    strAccessWhere = " WHERE waza = '" + comboBox1.Text + "'";
                    strAccessSelect = strAccessSelect + strAccessWhere;
                }
                else
                {
                    strAccessWhere = "";
                }

                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);

                myAccessConn.Open();

                OleDbDataReader reader = myAccessCommand.ExecuteReader();

                listBox1.Items.Clear();

                while (reader.Read())
                {
                    listBox1.Items.Add(reader.GetString(0) + "　" + reader.GetString(1) + "　" + reader.GetString(2));
                }
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed from the DataBase. " + ex.Message;
                return;
            }
            finally
            {
                label2.Text = strAccessSelect;
                myAccessConn.Close();
            }
        }

        //追加ボタンの処理

        private void button2_Click(object sender, EventArgs e)
        {
            strAccessInsert = "INSERT INTO waza(waza,tantou,timing)" + " VALUES('" + comboBox1.Text + "','" + comboBox2.Text + "','" + comboBox3.Text + "')";

            label2.Text = strAccessInsert;

            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed connection. " + ex.Message;
                return;
            }


            try
            {
                myAccessConn.Open();

                OleDbCommand cmd = new OleDbCommand(strAccessInsert, myAccessConn);
                cmd.ExecuteNonQuery();

                if (myAccessConn.State != ConnectionState.Closed)
                {
                    myAccessConn.Close();
                }
            }
            catch (OleDbException ex)
            {
                if (ex.ErrorCode == -2147467259)
                {
                    label1.Text = "データが重複しています";
                }
                else
                {
                    label1.Text = ex.Message;
                }
            }
            catch (Exception ex)
            {
                this.Width = 700; label1.Text = ex.Message;
            }
            finally
            {
                if (myAccessConn.State != ConnectionState.Closed)
                {
                    myAccessConn.Close();
                }
            }
        }

        //削除ボタンの処理

        private void button3_Click(object sender, EventArgs e)
        {
            strAccessDelete = "DELETE FROM waza" + " WHERE waza =  '" + comboBox1.Text + "'";

            label2.Text = strAccessDelete;

            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed connection. " + ex.Message;
                return;
            }

            try
            {
                myAccessConn.Open();

                OleDbCommand cmd = new OleDbCommand(strAccessDelete, myAccessConn);
                cmd.ExecuteNonQuery();

                if (myAccessConn.State != ConnectionState.Closed)
                {
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                this.Width = 700; label1.Text = ex.Message;
            }
            finally
            {
                if (myAccessConn.State != ConnectionState.Closed)
                {
                    myAccessConn.Close();
                }
            }
        }

        //更新ボタンの処理

        private void button4_Click(object sender, EventArgs e)
        {
            strAccessUpdate = "UPDATE waza" + " SET timing = " + comboBox3.Text + " WHERE waza =  '" + comboBox1.Text + "'";

            label2.Text = strAccessUpdate;

            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed connection. " + ex.Message;
                return;
            }

            try
            {
                myAccessConn.Open();

                OleDbCommand cmd = new OleDbCommand(strAccessUpdate, myAccessConn);
                cmd.ExecuteNonQuery();

                if (myAccessConn.State != ConnectionState.Closed)
                {
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                this.Width = 700; label1.Text = ex.Message;
            }
            finally
            {
                if (myAccessConn.State != ConnectionState.Closed)
                {
                    myAccessConn.Close();
                }
            }
        }

        //ラベルリセットの処理

        private void button5_Click(object sender, EventArgs e)
        {
            label1.ResetText();
            label2.ResetText();
        }

        //各ドロップダウンメニューの処理

        /* 演算子が足りないエラーが出てる。

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed connection. " + ex.Message;
                return;
            }

            try
            {
                if (comboBox1.Text.Length != 0)
                {
                    strAccessWhere = " WHERE waza = '" + comboBox1.Text + "'";
                    strAccessSelect = strAccessSelect + strAccessWhere;
                }
                else
                {
                    strAccessWhere = "";
                }

                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);

                myAccessConn.Open();

                OleDbDataReader reader = myAccessCommand.ExecuteReader();

                comboBox1.Items.Clear();

                while (reader.Read())
                {
                    comboBox1.Items.Add(reader.GetString(0));
                }
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed from the DataBase. " + ex.Message;
                return;
            }
            finally
            {
                label2.Text = strAccessSelect;
                myAccessConn.Close();
            }

        }

        private void comboBox2_DropDown(object sender, EventArgs e)
        {
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed connection. " + ex.Message;
                return;
            }

            try
            {
                if (comboBox2.Text.Length != 0)
                {
                    strAccessWhere = " WHERE tantou = '" + comboBox2.Text + "'";
                    strAccessSelect = strAccessSelect + strAccessWhere;
                }
                else
                {
                    strAccessWhere = "";
                }

                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);

                myAccessConn.Open();

                OleDbDataReader reader = myAccessCommand.ExecuteReader();

                comboBox2.Items.Clear();

                while (reader.Read())
                {
                    comboBox2.Items.Add(reader.GetString(1));
                }
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed from the DataBase. " + ex.Message;
                return;
            }
            finally
            {
                label2.Text = strAccessSelect;
                myAccessConn.Close();
            }

        }

        private void comboBox3_DropDown(object sender, EventArgs e)
        {
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed connection. " + ex.Message;
                return;
            }

            try
            {
                if (comboBox3.Text.Length != 0)
                {
                    strAccessWhere = " WHERE timing = '" + comboBox3.Text + "'";
                    strAccessSelect = strAccessSelect + strAccessWhere;
                }
                else
                {
                    strAccessWhere = "";
                }

                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);

                myAccessConn.Open();

                OleDbDataReader reader = myAccessCommand.ExecuteReader();

                comboBox3.Items.Clear();

                while (reader.Read())
                {
                    comboBox3.Items.Add(reader.GetString(2));
                }
            }
            catch (Exception ex)
            {
                label1.Text = "Error: Failed from the DataBase. " + ex.Message;
                return;
            }
            finally
            {
                label2.Text = strAccessSelect;
                myAccessConn.Close();
            }

        }*/
    }
}