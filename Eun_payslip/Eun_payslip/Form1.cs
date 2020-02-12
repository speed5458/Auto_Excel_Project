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
        String _User_list_File; //유저정보갖고있는 File의 주소저장
        List<string> _All_file = new List<string>(); //모든 선택된 파일의 경로를 갖고있다. 

        Dictionary<String, String> _User_Price = new Dictionary<string, string>(); //본인부담금 정보
        Dictionary<String, Info_data> _Result_Data = new Dictionary<string, Info_data>(); //장기요양 인정번호, 영수증번호, 급여 정보



        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 파일선택시 해당 파일의 수급자, 본인부담금데이터를 가져와 ListView에 뿌려준다.
        /// </summary>


        /// <summary>
        /// 저장된 File_Path 주소를 List에 뿌려준다.
        /// </summary>


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
            listView2.Columns[1].Width = -1; //컬럼크기를 텍스트 크기에 맞춰 정렬
            listView2.Columns[1].Width = -2; //컬럼크기를 텍스트 크기에 맞춰 정렬

        }

        private void Excel_Read_User()
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
                    arr[1] = ((range.Cells[j, 2] as Excel.Range).Value2.ToString() + " "); //User명
                    arr[2] = ((range.Cells[j, 8] as Excel.Range).Value2.ToString() + " "); //금액 컬럼 열 변경시 바꿔줘야한다.
                    itm = new ListViewItem(arr);
                    listView1.Items.Add(itm);
                    _User_Price.Add(arr[1], arr[2]);
                }
            }
            lblperiod.Text = (range.Cells[3, 5] as Excel.Range).Value2.ToString(); //급여제공기간
            //richTextBox1.Text = data;

            DeleteObject(worksheet1);
            DeleteObject(workbook);
            app.Quit();
            DeleteObject(app);
        }

        /// <summary>
        /// 개인파일을 열어 정보를 갖고온다. 1.성명, 2.장기요양 3.인정번호, 4.방문요양총금액
        /// </summary>
        private void Excel_Detail_Info_Read()
        {
            

            

            
            //for 문으로 리스트 크기만금 돌려서 모두 출력
            for(int i= 0; i < _All_file.Count; i++)
            {
                Excel.Application app = new Excel.Application();
                Excel.Workbook workbook = app.Workbooks.Open(Filename: _All_file[i]);
                Excel.Worksheet worksheet1 = workbook.Worksheets.Item[1]; //Index는 1부터 시작한다.
                app.Visible = false;
                Excel.Range range = worksheet1.UsedRange; //사용된 Range를 가져온다. Row와 Column의 정보를 갖고있다. 
                String Name = "";
                String Identify_No = "";
                String Price = "";
                Name = ((range.Cells[7, 2] as Excel.Range).Value2.ToString() + " "); //이름
                Identify_No = ((range.Cells[7, 16] as Excel.Range).Value2.ToString() + " "); //장기요양 인정번호

                for (int j = 21; j < range.Rows.Count; j++) //21은 시작하는열
                {
                    if ((string)(range.Cells[j, 28] as Excel.Range).Text.ToString() != "")
                    {
                        Price = ((range.Cells[j, 28] as Excel.Range).Value2.ToString() + " "); //마지막으로 입력된 값만 넣어준다. 나머지는 필요없음
                    }
                }

                Info_data info_data = new Info_data();
                info_data.Identify_No = Identify_No;
                info_data.Price = Price;

                _Result_Data.Add(Name, info_data);

                DeleteObject(worksheet1);
                DeleteObject(workbook);
                app.Quit();
                DeleteObject(app);
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

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            Excel_Detail_Info_Read();
        }

        /// <summary>
        /// 파일생성 클릭이 저장된 모든 데이터를 읽어서뿌려준다.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void simpleButton2_Click(object sender, EventArgs e)
        {
            //_User_Price
            foreach(KeyValuePair<String, Info_data> temp in _Result_Data)
            {
                Console.WriteLine("출력테스트 : 이름 = {0}, 인정번호 = {1}, 금액 = {2}", temp.Key, temp.Value.Identify_No, temp.Value.Price);
            }
        }
    }
    class Info_data
    {
        public String Identify_No { get; set; }
        public String Price { get; set; }
    }

}
