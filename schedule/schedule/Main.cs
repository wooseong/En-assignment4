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
            Console.SetWindowSize(190, 40); // 콘솔 창 사이즈 설정
            MainMenu mainMenu = new MainMenu(); // 메뉴 창 담고있는 변수

            mainMenu.ShowMenu(); // 메뉴 무한 루프

        }
    }
}