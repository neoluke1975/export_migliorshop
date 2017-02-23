using System;
using System.Windows.Forms;
using System.IO;
using FirebirdSql.Data.FirebirdClient;
using System.Diagnostics;


namespace export_migliorshop
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)

        {
            FbDataReader lettore = null;
            FbConnection conn = new FbConnection("User=SYSDBA;Password=masterkey;Database=" + textBox1.Text + ";DataSource=localhost");

            conn.Open();
            FbCommand query = new FbCommand("select a.kdes as Prodotto,"+
                                            "a.kean as EAN,"+
                                            "a.km10 as Minsan,"+
                                            "(select v_euro from vero_prezzo('today', a.km10,2)) as prezzo," +
                                            "max(l.e_costo_fd) as costo,"+
                                            "a.iva,"+
                                            "a.tip_revoca as revoca from costi_gr l inner join anapro a on a.km10 = l.km10 "+
                                            "where a.sud_merc='P' "+
                                            "group by  a.kdes,a.kean,a.km10,a.iva,a.tip_revoca "+
                                            "order by a.kdes", conn);
            
            StreamWriter scriviFile = new StreamWriter("c:/file_migliorshop/estratto.csv");
            lettore = query.ExecuteReader();
            while (lettore.Read())

            {
                string testa_descrizione ="PRODOTTO";
                string testa_ean = "EAN";
                string testa_minsan = "MINSAN";
                //string testa_disp = "Disp";
                string testa_prezzo ="PREZZO";
                string testa_costo = "COSTO";
                string testa_iva = "IVA";
                string testa_revoca = "REVOCA";

                scriviFile.WriteLine("{0,-41};{1,-13};{2,-13};{3,-8};{4,-8};{5,-3};{6,-6};", testa_descrizione, testa_ean,testa_minsan,testa_prezzo,testa_costo, testa_iva, testa_revoca);
                 while (lettore.Read())
                {
                    string descrizione = lettore.GetValue(0).ToString();
                    string ean = lettore.GetValue(1).ToString();
                    string minsan = lettore.GetValue(2).ToString();
                    //string disp = "0";
                    string iva = lettore.GetValue(5).ToString();
                    //sostituisco iva a 21 con 22 
                    if (iva == "21")
                    {
                        iva = "22";
                    }
                    if (iva == "20")
                    {
                        iva = "22";
                    }
                    var prezzo =double.Parse(lettore.GetValue(3).ToString());
                     
                    
                    var costo =double.Parse(lettore.GetValue(4).ToString());
                    //calcolo il costo se mancante
                    if (costo == 0.00)
                    {
                        costo = prezzo * (1 - 0.30);
                    }
                    //calcolo il prezzo se
                    if (prezzo == 0.00)
                    {
                        prezzo = costo / (1 - 0.30);
                    }
                    string revoca = lettore.GetValue(6).ToString();
                    if (costo==0.00)
                        if (prezzo == 0.00)
                        {
                            continue;

                        }

                    scriviFile.WriteLine("{0,-41};{1,13};{2,13};{3,8};{4,8};{5,3};{6,6};", descrizione, ean, minsan, prezzo.ToString("0.00"), costo.ToString("0.00"), iva, revoca);
                                      
                }
                MessageBox.Show("Export Completato!");
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start("c:/file_migliorshop/");
        }
    }
    }


