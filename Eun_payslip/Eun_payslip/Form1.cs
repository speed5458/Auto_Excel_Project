using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;

namespace Eun_payslip
{
    public partial class Form1 : Form
    {
        double _Process_Count = 1;

        int _status4 = 0;
        int _Result_status = 0;
        int _status2 = 0;
        int _status3 = 0;
        int _con = 0;

        String _document_template; //양식파일의 위치를 지정
        String _folder_path;
        String _User_list_File; //유저정보갖고있는 File의 주소저장
        List<string> _All_file = new List<string>(); //모든 선택된 파일의 경로를 갖고있다. 

        Dictionary<String, String> _User_Price = new Dictionary<string, string>(); //본인부담금 정보
        Dictionary<String, Info_data> _Result_Data = new Dictionary<string, Info_data>(); //장기요양 인정번호, 영수증번호, 급여 정보
        String _receipt_No;
        String _Period = "";

        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ofdFile.ShowDialog();
        }

        private void ofdFile_FileOk(object sender, CancelEventArgs e)
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();
            string filePath = ofdFile.FileName;
            _User_list_File = filePath; //파일의 풀경로 저장
            textBox1.Text = filePath.Split('\\')[filePath.Split('\\').Length - 1]; //파일이름만 읽어온다.

            listView1.Columns.Add("No.", 30);
            listView1.Columns.Add("수급자", 100);
            listView1.Columns.Add("본인부담금", 100);

            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;

            Excel_Read_User();

            _con = 1;

        }


        /// <summary>
        /// 다중파일선택버튼
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            ofdFile2.ShowDialog();
        }

        /// <summary>
        /// 파일 다중선택시 파일의 주소를 배열로 저장한다.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ofdFile2_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                listView2.Items.Clear();
                listView2.Columns.Clear();
                string[] arr = new string[2];
                ListViewItem itm;

                listView2.View = View.Details;
                listView2.GridLines = true;
                listView2.FullRowSelect = true;

                listView2.Columns.Add("No.", 50);
                listView2.Columns.Add("파일주소", 300);

                for (int i = 0; i < ofdFile2.FileNames.Length; i++)
                {
                    _All_file.Add(ofdFile2.FileNames[i].ToString()); //모든 파일의 경로를 List에 담는다.(나중에 빼서 하나씩열어서 데이터 저장하기위해서)
                    arr[0] = (i + 1).ToString(); //No. 저장
                    arr[1] = ofdFile2.FileNames[i].ToString(); //File 경로 저장
                    itm = new ListViewItem(arr);
                    listView2.Items.Add(itm);
                }

                label5.Text = _All_file.Count.ToString();
                listView2.Columns[1].Width = -1; //컬럼크기를 텍스트 크기에 맞춰 정렬
                listView2.Columns[1].Width = -2; //컬럼크기를 텍스트 크기에 맞춰 정렬

                _status2 = 3;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "ofdFile2_FileOk 에러");
            }
            


        }

        private void Excel_Read_User()
        {
            try
            {
                Excel.Application app = new Excel.Application();
                Excel.Workbook workbook = app.Workbooks.Open(Filename: _User_list_File);
                Excel.Worksheet worksheet1 = workbook.Worksheets.Item[1]; //Index는 1부터 시작한다.
                app.Visible = false;
                Excel.Range range = worksheet1.UsedRange; //사용된 Range를 가져온다. Row와 Column의 정보를 갖고있다. 



                //그리드에 정보를 담는다.
                string[] arr = new string[3];
                ListViewItem itm;

                for (int j = 10; j <= range.Rows.Count; j++)
                {
                    //String data_Result = (range.Cells[j, 2] as Excel.Range).Value2.ToString(); //null값이면 빠져나오기 위해
                    if ((string)(range.Cells[j, 2] as Excel.Range).Text.ToString() != "")
                    {
                        //data += ((range.Cells[j, 2] as Excel.Range).Value2.ToString() + " ");
                        arr[0] = (j - 9).ToString(); //No.
                        arr[1] = ((range.Cells[j, 2] as Excel.Range).Value2.ToString()); //User명
                        arr[2] = ((range.Cells[j, 8] as Excel.Range).Value2.ToString()); //금액 컬럼 열 변경시 바꿔줘야한다.
                        itm = new ListViewItem(arr);
                        listView1.Items.Add(itm);
                        _User_Price.Add(arr[1], arr[2]);
                    }
                }
                _Period = (range.Cells[3, 5] as Excel.Range).Value2.ToString(); //급여제공기간
                lblperiod.Text = _Period;
                //richTextBox1.Text = data;

                DeleteObject(worksheet1);
                DeleteObject(workbook);
                app.Quit();
                DeleteObject(app);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Excel_Read_User 에러");
            }
            
        }

        /// <summary>
        /// 개인파일을 열어 정보를 갖고온다. 1.성명, 2.장기요양 3.인정번호, 4.방문요양총금액
        /// </summary>
        private void Excel_Detail_Info_Read()
        {
            try
            {
                //for 문으로 리스트 크기만금 돌려서 모두 출력
                for (int i = 0; i < _All_file.Count; i++)
                {
                    Excel.Application app = new Excel.Application();
                    Excel.Workbook workbook = app.Workbooks.Open(Filename: _All_file[i]);
                    Excel.Worksheet worksheet1 = workbook.Worksheets.Item[1]; //Index는 1부터 시작한다.


                    app.Visible = false;
                    Excel.Range range = worksheet1.UsedRange; //사용된 Range를 가져온다. Row와 Column의 정보를 갖고있다. 
                    String Name = "";
                    String Identify_No = "";
                    String Price = "";
                    Name = ((range.Cells[7, 2] as Excel.Range).Value2.ToString()); //이름
                    Identify_No = ((range.Cells[7, 16] as Excel.Range).Value2.ToString()); //장기요양 인정번호

                    for (int j = 21; j < range.Rows.Count; j++) //21은 시작하는열
                    {
                        if ((string)(range.Cells[j, 28] as Excel.Range).Text.ToString() != "")
                        {
                            Price = ((range.Cells[j, 28] as Excel.Range).Value2.ToString()); //마지막으로 입력된 값만 넣어준다. 나머지는 필요없음
                        }
                    }

                    Info_data info_data = new Info_data();
                    info_data.Identify_No = Identify_No;
                    info_data.Price = Price;

                    _Result_Data.Add(Name, info_data);
                    int v = (int)((_Process_Count / _All_file.Count) * 100);
                    this.toolStripProgressBar1.Value = v;
                    this.toolStripStatusLabel2.Text = " " + v.ToString() + " %";
                    _Process_Count++;

                    DeleteObject(worksheet1);
                    DeleteObject(workbook);
                    app.Quit();
                    DeleteObject(app);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Excel_Detail_Info_Read 에러");
            }
            

        }
        private void ListView_Find()
        {
            try
            {
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    String listView_Name = this.listView1.Items[i].SubItems[1].Text;

                    if (_Result_Data.ContainsKey(listView_Name))
                    {
                        this.listView1.Items[i].BackColor = Color.LightGreen;
                    }
                    else
                    {
                        this.listView1.Items[i].BackColor = Color.Red;
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "ListView_Find 에러");
            }
            
        }

        /// <summary>
        /// 엑셀에 묶인 리소스지워준다.
        /// 엑셀에 묶인 리소스지워준다.
        /// </summary>
        /// <param name="obj"></param>
        private void DeleteObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("메모리 할당을 해제하는 중 문제가 발생하였습니다." + ex.ToString(), "경고!");
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// DataCheck 버튼, 다중선택된 파일 내용 읽어오고 해당 가져온 이름이 ListView1 리스트와 비교해서 체크해준다. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            Excel_Detail_Info_Read();
            ListView_Find();
            _Result_status = 3;
        }

        /// <summary>
        /// 출력버튼, 모든 정보를 읽어서, 파일로 변환한다.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void simpleButton2_Click(object sender, EventArgs e)
        {
            _Process_Count = 1;
            try
            {
                int num = int.Parse(textBox3.Text);
                foreach (KeyValuePair<String, String> temp in _User_Price)
                {
                    String Name = temp.Key; //이름
                    String User_Price = temp.Value.Trim(); //본인부담금
                    Info_data output = _Result_Data[Name.Trim()]; //이름을 넣고 값을 가져온다.

                    String Idenfify_No = output.Identify_No.Trim(); //인정번호
                    String Result_Price = output.Price.Trim(); //급여

                    String Year = textBox2.Text; //연도
                                                 //영수증번호를 입력받아서 넣어준다.

                    //String.Format("{0}-{1}", Year, num.ToString("000"));
                    String period = _Period.Trim().Substring(11, 23);

                    Console.WriteLine("이름 : {0} 본인부담금 : {1} 인정번호 : {2} 영수증번호 : {3} 급여 : {4}", Name, User_Price.Trim(), Idenfify_No, String.Format("{0}-{1}", Year, num.ToString("000")), Result_Price); ;
                    num++;

                    // 엑셀에 파일을 저장한다. 
                    Excel.Application excelApp = null;
                    excelApp = new Excel.Application();


                    Excel.Workbook wb = null;
                    Excel.Worksheet ws = null;
                    try
                    {
                        wb = excelApp.Workbooks.Open(_document_template);
                        ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
                        ws.Range["A8"].Value = Name; //이름
                        ws.Range["C8"].Value = Idenfify_No; //장기요양 인정번호
                        ws.Range["D8"].Value = period; //급여제공기간
                        ws.Range["F8"].Value = String.Format("{0}-{1}", Year, num.ToString("0000")); ; //영수증번호
                        ws.Range["D10"].Value = User_Price; //본인부담금
                        ws.Range["D12"].Value = Result_Price; //급여 계
                        wb.SaveAs(_folder_path + "\\" + Name + ".xlsx");

                        int v = (int)((_Process_Count / _All_file.Count) * 100);
                        this.toolStripProgressBar1.Value = v;
                        this.toolStripStatusLabel2.Text = " " + v.ToString() + " %";
                        _Process_Count++;
                    }
                    catch (Exception ex)
                    {
                        if (wb != null) wb.Close();
                        if (excelApp != null) excelApp.Quit();
                        return;
                    }

                    if (wb != null) wb.Close();

                }
                
                
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "simpleButton2_Click 에러");
            }
            


            ///////////////////출력을합니다.
        }
        #region
        /// <summary>
        /// Timer 모음, 버튼 색바뀌는 Timer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (_status2 == 0)
            {
                _status2 = 1;
                button3.BackColor = Color.Red;
            }
            else if (_status2 == 1)
            {
                _status2 = 0;
                button3.BackColor = Color.White;
            }
            else if (_status2 == 3)
            {
                _status2 = 4;
                button3.BackColor = Color.LightGreen;
            }
            else if (_status2 == 4)
            {
                _status2 = 3;
                button3.BackColor = Color.White;
            }
        }
        

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (_Result_status == 0)
            {
                _Result_status = 1;
                button4.BackColor = Color.Red;
            }
            else if (_Result_status == 1)
            {
                _Result_status = 0;
                button4.BackColor = Color.White;
            }
            else if (_Result_status == 3 && _con == 1)
            {
                _Result_status = 4;
                button4.BackColor = Color.LightGreen;
            }
            else if (_Result_status == 4 && _con == 1)
            {
                _Result_status = 3;
                button4.BackColor = Color.White;
            }
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            if (_status3 == 0)
            {
                _status3 = 1;
                btnStatus.BackColor = Color.Red;
            }
            else if (_status3 == 1)
            {
                _status3 = 0;
                btnStatus.BackColor = Color.White;
            }
            else if (_status3 == 3)
            {
                _status3 = 4;
                btnStatus.BackColor = Color.LightGreen;
            }
            else if (_status3 == 4)
            {
                _status3 = 3;
                btnStatus.BackColor = Color.White;
            }
        }
        private void timer4_Tick(object sender, EventArgs e)
        {
            if (_status4 == 0)
            {
                _status4 = 1;
                button5.BackColor = Color.Red;
            }
            else if (_status4 == 1)
            {
                _status4 = 0;
                button5.BackColor = Color.White;
            }
            else if (_status4 == 3)
            {
                _status4 = 4;
                button5.BackColor = Color.LightGreen;
            }
            else if (_status4 == 4)
            {
                _status4 = 3;
                button5.BackColor = Color.White;
            }
        }
        #endregion

        class Info_data
        {
            public String Identify_No { get; set; }
            public String Price { get; set; }
        }

        private void btnStatus_Click(object sender, EventArgs e)
        {

            fbDialog.ShowDialog();
            _folder_path = fbDialog.SelectedPath;
            _status3 = 3;

        }

        /// <summary>
        /// 양식파일위치 지정
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            ofdFile4.ShowDialog();
        }

        private void ofdFile4_FileOk(object sender, CancelEventArgs e)
        {
            _document_template = ofdFile4.FileName;
            _status4 = 3;

        }

        private void button6_Click(object sender, EventArgs e)
        {
            HelpForm help = new HelpForm();
            help.ShowDialog();
        }
    }
}
