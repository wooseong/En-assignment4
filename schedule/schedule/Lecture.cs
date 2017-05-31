using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;


namespace schedule
{
    public class Lecture
    {
        #region 변수선언
        Excel.Application excelApplication;
        Excel.Workbook excelWorkbook;
        Excel._Worksheet excelWorkSheet;
        Excel.Range excelRange;
        int rowCount; // 시간표 행 개수
        int colCount; // 시간표 열 개수
        private string searchWithNumberString; // SearchLecturePrintWith의 메뉴 번호 선택 변수
        private int searchWithNumber; // SearchLecturePrintWith의 메뉴 번호에 따른 엑셀 열 번호
        private string searchWitinformation; //  SearchLecturePrintWith의 새부 검색 내용 변수
        private int check; //새부 검색 내용이 출력 되었는지 확인
        private string subjectNumber;
        private string subjectclass;

        List<string> selectLecture = new List<string>(new string[] { "", }); //
        List<string> selectIntrestLecture = new List<string>(new string[] { "", });
        List<string> searchWithList = new List<string>(new string[]
        {"뒤로가기",
        "개설학과전공",
        "교과목명",
        "이수 구분",
        "학년",
        "학점",
        "요일 및 시간",
        "교수명",
        "강의언어",
        "전체"
        });
        #endregion

        public Lecture(string directory)// 엑셀 불러오고 각 변수에 초기화
        {
            excelApplication = new Excel.Application(); // Excel 첫번째 워크시트 가져오기
            excelWorkbook = excelApplication.Workbooks.Open(directory);
            excelWorkSheet = excelWorkbook.Sheets[1];
            excelRange = excelWorkSheet.UsedRange;

            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;
        }
        ~Lecture()
        {

            GC.Collect();
            GC.WaitForPendingFinalizers();
            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(excelWorkSheet);

            //close and release
            excelWorkbook.Close();
            Marshal.ReleaseComObject(excelWorkbook);

            //quit and release
            excelApplication.Quit();
            Marshal.ReleaseComObject(excelApplication);
        }

        public bool SearchLectureWith() // 강의 출력을 위한 세부사항 설정
        {
            do// 어떤 세부 목록으로 검색할지 번호 선택
            {
                Console.Clear();
                Console.WriteLine("\n\n-------------------------------------------------------------------강의 출력-------------------------------------------------------------------");
                Console.WriteLine("\n\n\n");
                for (int i = 0; i < 10; i++)
                    Console.WriteLine("\t\t\t\t\t\t\t\t{0}. {1}", i, searchWithList[i]);
                Console.Write("\t\t\t\t\t\t\t\t어떤 것을 통해 보시겠습니까?  ");
                searchWithNumberString = Console.ReadLine();
            } while (!(searchWithNumberString.Equals("1") || searchWithNumberString.Equals("2") || searchWithNumberString.Equals("3") ||
              searchWithNumberString.Equals("4") || searchWithNumberString.Equals("5") || searchWithNumberString.Equals("6") ||
              searchWithNumberString.Equals("7") || searchWithNumberString.Equals("8") || searchWithNumberString.Equals("9") || searchWithNumberString.Equals("0")));// 어떤 세부 목록으로 검색할지 번호 선택
            if (searchWithNumberString.Equals("0")) return false;// 0입력시, 무한루프 나가기(뒤로가기)

            check = -1; // do while문이 처음인지 위해 -1로 마음대로 의미지정
            do
            {
                #region // 어떤 세부 목록으로 검색할지 내용 입력
                if (check != -1) // 첫 do while이 아닌 경우(0~9이외의 값을 입력시) 메뉴를 다시 출력해야하므로
                {
                    Console.Clear();
                    Console.WriteLine("--------------------------------------------------------------------------------강의 출력--------------------------------------------------------------------------------");
                    Console.WriteLine("\n\n\n\n\n");
                    for (int i = 0; i < 10; i++)
                        Console.WriteLine("\t\t\t\t\t\t\t\t{0}. {1}", i, searchWithList[i]);
                    Console.WriteLine("\t\t\t\t\t\t\t\t어떤 것을 통해 보시겠습니까?  {0}", searchWithNumberString);
                }// 첫 do while이 아닌 경우(0~9이외의 값을 입력시) 메뉴를 다시 출력해야하므로
                Console.Write("\t\t\t\t\t\t\t\t{0}  ", searchWithList[Convert.ToInt32(searchWithNumberString)]);
                searchWitinformation = Console.ReadLine();
                if (searchWitinformation.Equals("0")) return false;// 0입력시, 무한루프 나가기(뒤로가기)
                check = 0; // 출력된 강의가 몇개인지 count하기 위해 초기화
                #endregion

                #region 메뉴 조건문
                if (searchWithNumberString.Equals("1"))// 개설학과
                {
                    searchWithNumber = 2;
                    SearchLecturePrint();
                }
                else if (searchWithNumberString.Equals("2"))// 교과목명
                {
                    searchWithNumber = 5;
                    SearchLecturePrint();
                }
                else if (searchWithNumberString.Equals("3"))// 이수 구분
                {
                    searchWithNumber = 6;
                    SearchLecturePrint();
                }
                else if (searchWithNumberString.Equals("4"))// 학년
                {
                    searchWithNumber = 7;
                    SearchLecturePrint();
                }
                else if (searchWithNumberString.Equals("5"))// 학점
                {
                    searchWithNumber = 8;
                    SearchLecturePrint();
                }
                else if (searchWithNumberString.Equals("6"))// 요일 및 시간
                {
                    searchWithNumber = 9;
                    SearchLecturePrint();
                }
                else if (searchWithNumberString.Equals("7"))// 교수명
                {
                    searchWithNumber = 11;
                    SearchLecturePrint();
                }
                else if (searchWithNumberString.Equals("8"))// 강의언어
                {
                     
                    searchWithNumber = 12;
                    SearchLecturePrint();
                }
                else if (searchWithNumberString.Equals("9"))// 전체
                {
                    searchWithNumber = 0;
                    SearchLecturePrint();
                }
                #endregion
            } while (check == 0);
            return true; // 무의미한 리턴
        }
        public void SearchLecturePrint()
        {
            for (int i = 1; i <= rowCount; i++)
            {

                if (excelRange.Cells[i, searchWithNumber].Value2.ToString() != searchWitinformation)
                    continue;

                for (int j = 1; j <= colCount; j++)
                {
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null) // 메뉴 번호에 따른 세부내용 확인
                        Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t ");
                    check++;

                }
                Console.WriteLine("");

            }
            if (check != 0)
                Console.ReadLine();
        }
        public List<string> AddLecture(int menuNumber)
        {
            do
            {
                Console.Clear();
                Console.Write("\t수강 신청하실 과목의 학수번호와 분반 입력하세요.  \n\t학수 번호  :  ");
                subjectNumber = Console.ReadLine();
                Console.Write("\t분     반  :  ");
                subjectclass = Console.ReadLine();
                check = 0;
                for (int i = 1; i <= rowCount; i++)
                {

                    if ((excelRange.Cells[i, 3].Value2.ToString() == subjectNumber) && (excelRange.Cells[i, 4].Value2.ToString() == subjectclass))
                    {
                        check++;
                        for (int j = 1; j < colCount; j++)
                        {
                            if (menuNumber == 2)
                                selectLecture.Add(excelRange.Cells[i, j].Value2.ToString());
                            else if (menuNumber == 5)
                                selectIntrestLecture.Add(excelRange.Cells[i, j].Value2.ToString());
                        }
                        Console.Write("정상적으로 추가되셨습니다.");
                        Console.ReadLine();

                    }
                }
                if (check == 0)
                {
                    Console.Write("그런 수업은 없습니다.");
                    Thread.Sleep(1000);
                }
                //else
                //{
                //    for (int j = 0; j < colCount; j++)
                //    {
                //        Console.Write("{0,10}", selectLecture[j]);

                //    }
                //   // Console.WriteLine(selectLecture);
                //    Console.ReadLine();
                //}
            } while (check == 0);
            return selectLecture;
        }
        public void ErasureLecture(int menuNumber)
        {
            Console.Clear();
            Console.Write("\t 삭제하실 과목의 학수번호와 분반 입력하세요.  \n\t학수 번호  :  ");
            subjectNumber = Console.ReadLine();
            Console.Write("\t분     반  :  ");
            subjectclass = Console.ReadLine();
            check = 1;
            for (int i = 1; i <= rowCount; i++)
            {

                if ((excelRange.Cells[i, 3].Value2.ToString() == subjectNumber) && (excelRange.Cells[i, 4].Value2.ToString() == subjectclass))
                {
                    check++;
                    for (int j = 1; j < colCount; j++)
                    {
                        if (menuNumber == 2)
                            selectLecture.Remove(excelRange.Cells[i, j].Value2.ToString());
                        else if (menuNumber == 5)
                            selectIntrestLecture.Remove(excelRange.Cells[i, j].Value2.ToString());

                    }
                    Console.Write("정상적으로 삭제되었습니다.");
                    Console.ReadLine();

                }
            }
            if (check == 0)
            {
                Console.Write("그런 수업은 없습니다.");
                Thread.Sleep(1000);
            }
            else
            {
                for (int j = 0; j < colCount; j++)
                {
                    Console.Write("{0,10}", selectLecture[j]);

                }
                // Console.WriteLine(selectLecture);
                Console.ReadLine();
            }
        }
    }
}
