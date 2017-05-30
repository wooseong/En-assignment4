using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using schedule;

namespace Sandbox
{
    public class Read_From_Excel
    {
        public static void Main()
        {
            Console.SetWindowSize(180, 50); // 콘솔 창 사이즈 설정

            FindLecture findLecture = new FindLecture(@"C:\Users\신우성\Desktop\excelStudy.xlsx"); // 엑셀 파일 위치
            findLecture.initExcel(); // 엑셀을 열고 사용할 준비 단계

            MainMenu mainMenu = new MainMenu();
            while (mainMenu.ShowMenu()) ;

            //findLecture.searchLecture("컴퓨터공학과");
            //findLecture.searchLecture("디지털공학과");



        }
    }
}