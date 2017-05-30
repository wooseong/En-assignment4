using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace schedule
{
    class MainMenu
    {
        private string menuNumber;
        public bool ShowMenu()
        {
            Console.Clear();
            Console.WriteLine("\n\n\n\n\n\n");
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t 1.강의 출력");
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t 2.강의 추가");
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t 3.강의 삭제");
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t 4.관심 과목 추가");
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t 5.관심 과목 삭제");
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t 6.관심 과목 출력");
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t 7.시간표 출력");
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t 0. exit");
            Console.Write("\n\t\t\t\t\t\t\t\t\t\t\t");
            menuNumber = Console.ReadLine();

            if (menuNumber.Equals("0")) return false; // 끝내야해서 false
            else if (menuNumber.Equals("1"))
            {
                return true;
            }
            else if (menuNumber.Equals("2"))
            {
                return true;
            }
            else if (menuNumber.Equals("3"))
            {
                return true;
            }
            else if (menuNumber.Equals("4"))
            {
                return true;
            }
            else if (menuNumber.Equals("5"))
            {
                return true;
            }
            else if (menuNumber.Equals("6"))
            {
                return true;
            }
            else if (menuNumber.Equals("7"))
            {
                return true;
            }
            else
            {
                Console.WriteLine("\t\t\t\t\t\t\t\t\t\t잘못입력하셨습니다.");
                Console.Write("\t\t\t\t\t\t\t\t\t\t 다시 입력하세요.");
                Thread.Sleep(1000);
                return true;
            }
        }
    }
}
