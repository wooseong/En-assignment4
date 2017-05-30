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
    public class FindLecture
    {
        #region 변수선언
        Excel.Application excelApplication;
        Excel.Workbook excelWorkbook;
        Excel._Worksheet excelWorkSheet;
        Excel.Range excelRange;
        int rowCount; // 시간표 행 개수
        int colCount; // 시간표 열 개수
        private string searchWithNumberString; // searchLectureWith의 메뉴 번호 선택 변수
        private int searchWithNumber; // searchLectureWith의 메뉴 번호에 따른 엑셀 열 번호
        private string searchWitinformation; //  searchLectureWith의 새부 검색 내용 변수
        private int check; //새부 검색 내용이 출력 되었는지 확인

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

        public FindLecture(string directory)// 엑셀 불러오고 각 변수에 초기화
        {
            excelApplication = new Excel.Application(); // Excel 첫번째 워크시트 가져오기
            excelWorkbook = excelApplication.Workbooks.Open(directory);
            excelWorkSheet = excelWorkbook.Sheets[1];
            excelRange = excelWorkSheet.UsedRange;

            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;
        }
        ~FindLecture()
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

        public void SearchLectureWith()
        {
            do
            {
                Console.Clear();
                Console.WriteLine("--------------------------------------------------------------------------------강의 출력--------------------------------------------------------------------------------");
                Console.WriteLine("\n\n\n\n\n");
                for (int i = 0; i < 10; i++)
                    Console.WriteLine("\t\t\t\t\t\t\t\t\t\t{0}. {1}", i, searchWithList[i]);
                Console.Write("\t\t\t\t\t\t\t\t\t\t어떤 것을 통해 보시겠습니까?  ");
                searchWithNumberString = Console.ReadLine();
            } while (!(searchWithNumberString.Equals("1") || searchWithNumberString.Equals("2") || searchWithNumberString.Equals("3") ||
              searchWithNumberString.Equals("4") || searchWithNumberString.Equals("5") || searchWithNumberString.Equals("6") ||
              searchWithNumberString.Equals("7") || searchWithNumberString.Equals("8") || searchWithNumberString.Equals("9") || searchWithNumberString.Equals("0")));

            check = -1; // do while문이 처음인지 위해 -1로 마음대로 의미지정
            do
            {if (check != -1)
                {
                    Console.Clear();
                    Console.WriteLine("--------------------------------------------------------------------------------강의 출력--------------------------------------------------------------------------------");
                    Console.WriteLine("\n\n\n\n\n");
                    for (int i = 0; i < 10; i++)
                        Console.WriteLine("\t\t\t\t\t\t\t\t\t\t{0}. {1}", i, searchWithList[i]);
                    Console.WriteLine("\t\t\t\t\t\t\t\t\t\t어떤 것을 통해 보시겠습니까?  {0}", searchWithNumberString);
                }
                Console.Write("\t\t\t\t\t\t\t\t\t\t{0}  ", searchWithList[Convert.ToInt32(searchWithNumberString)]);
                searchWitinformation = Console.ReadLine();
                check = 0;

                if (searchWithNumberString.Equals("0")) { check = -1; }// 무한루프 나가기 위한 값, 값은 무의미
                else if (searchWithNumberString.Equals("1"))// 개설학과
                {
                    searchWithNumber = 2;
                    SearchLecture();
                }
                else if (searchWithNumberString.Equals("2"))// 교과목명
                {
                    searchWithNumber = 5;
                    SearchLecture();
                }
                else if (searchWithNumberString.Equals("3"))// 이수 구분
                {
                    searchWithNumber = 6;
                    SearchLecture();
                }
                else if (searchWithNumberString.Equals("4"))// 학년
                {
                    searchWithNumber = 7;
                    SearchLecture();
                }
                else if (searchWithNumberString.Equals("5"))// 학점
                {
                    searchWithNumber = 8;
                    SearchLecture();
                }
                else if (searchWithNumberString.Equals("6"))// 요일 및 시간
                {
                    searchWithNumber = 9;
                    SearchLecture();
                }
                else if (searchWithNumberString.Equals("7"))// 교수명
                {
                    searchWithNumber = 11;
                    SearchLecture();
                }
                else if (searchWithNumberString.Equals("8"))// 강의언어
                {
                    searchWithNumber = 12;
                    SearchLecture();
                }
                else if (searchWithNumberString.Equals("9"))// 전체
                {
                    searchWithNumber = 0;
                    SearchLecture();
                }
            } while (check == 0);
        }
        public void SearchLecture()
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
                Console.WriteLine("\r");

            }
            if (check != 0)
                Console.ReadLine();
        }
    }
}
