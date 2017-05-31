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
        Lecture Lecture = new Lecture(@"C:\Users\신우성\Desktop\excelStudy.xlsx"); // 엑셀 파일 위치
        TimeSheet sheet = new TimeSheet();
        private string menuNumber;

        public bool ShowMenu()
        {
            #region 메뉴 출력문
            Console.Clear();
            Console.WriteLine("\n\n-------------------------------------------------------------------Main Menu-------------------------------------------------------------------");
            Console.WriteLine("\n\n\n");
            Console.WriteLine("\t\t\t\t\t\t\t\t 1.강의 출력"); // 엑셀 -> 콜솔로 출력
            Console.WriteLine("\t\t\t\t\t\t\t\t 2.수강 신청"); // 엑셀에 추가
            Console.WriteLine("\t\t\t\t\t\t\t\t 3.수강 철회");
            Console.WriteLine("\t\t\t\t\t\t\t\t 4.관심 과목 출력");
            Console.WriteLine("\t\t\t\t\t\t\t\t 5.관심 과목 추가");
            Console.WriteLine("\t\t\t\t\t\t\t\t 6.관심 과목 삭제");
            Console.WriteLine("\t\t\t\t\t\t\t\t 7.시간표 출력"); // 콘솔 -> 엑셀로
            Console.WriteLine("\t\t\t\t\t\t\t\t 0. exit");
            Console.Write("\n\t\t\t\t\t\t\t\t\t");
            #endregion
            menuNumber = Console.ReadLine();

            #region 메뉴 조건문
            if (menuNumber.Equals("0")) return false; // 끝내야해서 false
            else if (menuNumber.Equals("1"))// 강의 출력
            {
                Lecture.SearchLectureWith();
                return true;
            }
            else if (menuNumber.Equals("2"))// 수강 신청
            {
                Lecture.AddLecture(2);
                return true;
            }
            else if (menuNumber.Equals("3"))// 수강 철회
            {
                Lecture.ErasureLecture(2);
                return true;
            }
            else if (menuNumber.Equals("4"))// 관심 과목 출력
            {
                return true;
            }
            else if (menuNumber.Equals("5"))// 관심 과목 추가
            {
                Lecture.AddLecture(5);
                return true;
            }
            else if (menuNumber.Equals("6"))// 관심 과목 삭제
            {
                Lecture.ErasureLecture(5);
                return true;
            }
            else if (menuNumber.Equals("7"))// 시간표 출력
            {
                sheet.printTimeSheet();
                Console.ReadLine();
                return true;
            }
            else
            {
                Console.WriteLine("\t\t\t\t\t\t\t\t잘못입력하셨습니다.");
                Console.Write("\t\t\t\t\t\t\t\t 다시 입력하세요.");
                Thread.Sleep(1000);
                return true;
            }
            #endregion
        }
    }
}
