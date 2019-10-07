using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Testy1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
       // public class RegExFolders
       // {
       //     public static List<string> ListFIles(string parentFolder)
       //     {
       //
       //         Regex rx = new Regex(@"[A-Z]{2}");
       //         var result = new List<string>();
       //         var di = new DirectoryInfo(parentFolder);
       //         var folders = di.EnumerateDirectories().ToList().Where(d => rx.IsMatch(d.Name)).ToList();
       //         folders.ForEach(f => result.AddRange(Directory.GetFiles(f.FullName, "*.pdf")));
       //         return result;
       //     }

       // }

        

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
            panel1.BackColor = Color.LightGray;
        }

        
        string[] xmls;
        public void panel1_DragDrop(object sender, DragEventArgs e)
        {
            

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            string fldr = new DirectoryInfo(files[0]).Name;
                     

           
       
            if (fldr != "10_XMLs_org")
            {
                MessageBox.Show("Drag '10_XMLs_org' folder!");
                panel1.BackColor = Color.White;

                
            } else
            {
                xmls = Directory.GetFiles(files[0], "*.xml");
                textBox1.Text = files[0];
                textBox2.Text = xmls.Length.ToString();
                xmlPath = files[0];
                panel1.BackColor = Color.Gray;

                panel2.Visible = true;
                label3.Visible = true;
                label5.Visible = true;
                textBox3.Visible = true;
                label4.Enabled = false;
                label6.Visible = true;
            }
            

        }

        string xmlPath;

        private void panel2_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
            panel2.BackColor = Color.LightGray;
        }
        public void panel2_DragDrop(object sender, DragEventArgs e)
        {

            string FolderColor(string num)
            {
                if(num == "16777215")
                {
                  return "SBD_White";
                } else if(num == "5296274")
                {
                    return "SBD_Green";
                }
                else if (num == "65535")
                {
                    return "SBD_Yellow";
                }
                else if (num == "255")
                {
                    return "SBD_Red";

                } else if (num == "192")
                {
                    return "SBD_Dark_Red";

                } else if (num == "49407")
                {
                    return "SBD_Orange";
                }
                else if (num == "5287936")
                {
                    return "SBD_Dark_Green";
                }
                else if (num == "15773696")
                {
                    return "SBD_Light_Blue";
                }
                else if (num == "12611584")
                {
                    return "SBD_Blue";
                }
                else if (num == "6299648")
                {
                    return "SBD_Dark_Blue";
                }
                else if (num == "10498160")
                {
                    return "SBD_Purple";
                }


                else
                {

                    var unknownColor = string.Concat("SBD_" + num);
                    MessageBox.Show("Custom color! Folder named: " + unknownColor);
                    return unknownColor;
                }


                
            }

            if (xmlPath == null)
            {
                MessageBox.Show("Drag '10_XMLs_org' folder first");
                panel2.BackColor = Color.White;
            }
            else
            {

                string[] exclefile = (string[])e.Data.GetData(DataFormats.FileDrop, false);

                if (exclefile[0].IndexOf(".xlsx") == -1)
                {
                    MessageBox.Show("Must be an XLSX file");
                    panel2.BackColor = Color.White;
                }
                else
                {
                    panel2.BackColor = Color.Gray;
                    label7.Visible = true;
                    // else

                    textBox3.Text = exclefile[0];
                    label5.Enabled = false;


                    string fle = exclefile[0].ToString();
                    //MessageBox.Show(fle);
                    Excel excel = new Excel(fle);
                    //wyzej dodaj 2gi param ewentualnie
                    //MessageBox.Show(excel.ReadCell(13, 2));
                    // MessageBox.Show(excel.ReadColor(13, 2));
                    //MessageBox.Show(excel.Get1stRowLength().ToString());

                    int exlLeng = Int32.Parse(excel.Get1stRowLength());

                    for (int i = 1; i < exlLeng; i++)

                        {
                        //MessageBox.Show(excel.ReadCellNum(i, 1).ToString());
                        l.Add(excel.ReadCellNum(i, 1));


                    }

                    l = l.Distinct().ToList();
                    l.Remove(0);
                    //comboBox1.DataSource = l;

                    //customMSG.Show(l);
                    //MessageBox.Show("wybrano batch (ze zmiennej)" + BaczNumCustom);
                    int copyCounter = 0;

                    using (customMSG form22 = new customMSG())
                    {
                        form22.ComboCustom.DataSource = l;
                        if (form22.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {

                            selekt = Convert.ToInt32(form22.ComboCustom.SelectedItem);
                            //MessageBox.Show("Batch: " + selekt);

                            double skan;
                            string curColor;
                            string procka;
                            string prockaSearch;
                            string SBD_Color;


                            DirectoryInfo D1 = Directory.CreateDirectory(xmlPath);
                            string parent = Directory.GetParent(xmlPath).FullName;
                            //MessageBox.Show(parent);
                            DirectoryInfo D2 = Directory.CreateDirectory(parent);
                            

                            for (int i = 1; i < exlLeng; i++)
                            {
                                skan = excel.ReadCellNum(i, 1);
                                if (skan != selekt)
                                {
                                continue;
                                }
                                else
                                {
                                    //MessageBox.Show("batcz " + selekt + " wystepuje w komorce " + i + " procka: " + excel.ReadCell(i, 2) + " kolorek: " + excel.ReadColor(i, 2) + " w folderze " + xmlPath);
                                    curColor = excel.ReadColor(i, 2);
                                    procka = excel.ReadCell(i, 2);
                                    prockaSearch = String.Concat(procka + "*");
                                    string[] xmlList = Directory.GetFiles(xmlPath, prockaSearch);

                                    //pobiera pelna nazwe folderu SBD_xxxxx
                                    SBD_Color = FolderColor(curColor);

                                    D2.CreateSubdirectory(SBD_Color);
                                    
                                    string newSBDfolderFullPath = Path.Combine(xmlPath, SBD_Color);
                                    
                                    DirectoryInfo D3 = Directory.CreateDirectory(newSBDfolderFullPath);
                                    MessageBox.Show("dddddd");
                                    string the10xmlorg = "10_XMLs_org";
                                    D3.CreateSubdirectory(the10xmlorg);

                                    foreach (string f in xmlList)
                                    {
                                        string fName = f.Substring(xmlPath.Length + 1);

                                        // Use the Path.Combine method to safely append the file name to the path.
                                        // Will overwrite if the destination file already exists.
                                        File.Copy(Path.Combine(xmlPath, fName), Path.Combine(Path.Combine(parent, SBD_Color), fName), true);
                                        copyCounter += 1;
                                    }

                                }
                            }
                        }
                    }
                    if (copyCounter == xmls.Length)
                    {
                        MessageBox.Show("Copied: " + copyCounter.ToString() + " / " + xmls.Length.ToString() + "\n" + "\n" + "OK!");
                    }
                    else
                    {
                        MessageBox.Show("ERROR! Copied " + copyCounter.ToString() + " / " + xmls.Length.ToString() + "\n" + "\n" + "Did you choose right batch number?");
                    }
                    excel.xClose();

                } // else in
            } // else
        }


       


        int selekt;


        class Excel
        {
            string exPath = "";
            _Application excel = new _Excel.Application();
            Workbook wb;
            Worksheet ws;

            //public Excel(string exPath, int exSheet)
            public Excel(string exPath)
            {
                this.exPath = exPath;
                wb = excel.Workbooks.Open(exPath);
                //ws = wb.Worksheets[exSheet];
                ws = wb.ActiveSheet;
            }
            public string ReadColor(int i, int j)
            {
                string txtColor = ws.Cells[i, j].Interior.Color.ToString();
                
                 return txtColor; 

            }
            public string Get1stRowLength()
            {
                long lastRow;
                long fullRow;

                fullRow = ws.Rows.Count;
                //int lastUsedRow = ws.Cells.SpecialCells(excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
                return lastRow.ToString();
            }

            public string ReadCell(int i, int j)
            {


                //i++;
                //j++;

                if (ws.Cells[i, j].Value2 != null)
                {
                    return ws.Cells[i, j].Value2;
                }
                
                else
                {
                    return "null";
                }
            }

            public double ReadCellNum(int i, int j)
            {


                //i++;
                //j++;

                if (ws.Cells[i, j].Value2 != null)
                {
                    if(ws.Cells[i, j].Value2 is string)
                    {
                        return 0;
                    }
                    else
                    {
                        return ws.Cells[i, j].Value2;
                    }
                    
                }

                else
                {
                    return 0;
                }
            }

            




            public void xClose()
            {
                wb.Close();
            }

            public void niekumam()
            {
                

                
            }


        }

        

       

        List<double> l = new List<double>();
        private void button3_Click(object sender, EventArgs e)
        {
             
            //double val = textBox4.Text;
            //MessageBox.Show(val);
            
            //if (!l.Exists(x => x == val))
            //l.Add(val);

            
            l = l.Distinct().ToList();

            
        }

        
        

    }
}
