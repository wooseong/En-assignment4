using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.IO;

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
        private int searchWithNumber; // SearchLecturePrintWith의 메뉴 번호에 따른 엑셀 열 번호 단, 0이면 전체 출력
        private string searchWitinformation; //  SearchLecturePrintWith의 새부 검색 내용 변수
        private int check; //새부 검색 내용이 출력 되었는지 확인
        List<LectureVO> lecture = new List<LectureVO>();

        private string subjectNumber;
        private string subjectclass;

        List<string> selectLecture = new List<string>(new string[] { "", });
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
        }); // 세부 목록 검색 무엇을 할지 목록
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

        public void LectureSet() // 강의를 엑셀로부터 불러와서 초기화하기
        {
            string temporarily;
            //string nameTemporarily;
            string week = "";
            string time = "";
            string lectureRoom = "";// 강의실 10
            string professorName = ""; // 교수명 11
            List<string> temporarilyList = new List<string>();

            for (int i = 2; i <= rowCount; i++)// 엑셀 행 돌리기
            {
                week = "";
                time = "";
                lectureRoom = "";
                professorName = "";
                temporarily = excelRange.Cells[i, 9].Value2.ToString();

                for (int j = 1; j <= colCount; j++)// 엑셀 엘 돌리기
                {
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                    {
                        switch (j)// 열 번호에 따라 정하기 전에 조건 절차를 위하 나누기
                        {

                            case 9: // 엑셀에서 열 기준으로 9번째인 요일 및 강의 시간
                                for (int k = 0; k < temporarily.Length; k++)
                                {
                                    switch (temporarily[k]) //화목13:30-15:00,화18:00-20:00 이런거 나누기
                                    {
                                        case '월':
                                        case '화':
                                        case '수':
                                        case '목':
                                        case '금':
                                            week = week + temporarily[k];
                                            break;
                                        default:
                                            time = time + temporarily[k];
                                            break;
                                    }
                                }
                                break;
                            case 10: // 10번째 강의실
                                if (excelRange.Cells[i, 10].Value2.ToString().ToString().Length < 2)
                                {
                                    lectureRoom = "\t\t";
                                }
                                else
                                    lectureRoom = excelRange.Cells[i, 10].Value2.ToString();
                                break;
                            case 11: // 11번째 교수명
                                if (excelRange.Cells[i, 11].Value2.ToString() == null)
                                {
                                    professorName = "\t\t";
                                }
                                else
                                    professorName = excelRange.Cells[i, 11].Value2.ToString();
                                break;
                            default: // 특별한 경우 없이 string 으로 넣어도 되는 나머지
                                temporarilyList.Add(excelRange.Cells[i, j].Value2.ToString());
                                break;
                        }
                    }
                }
                lecture.Add(new LectureVO(Int32.Parse(temporarilyList[0]), temporarilyList[1], temporarilyList[2],
                     temporarilyList[3], temporarilyList[4], temporarilyList[5],
                      Int32.Parse(temporarilyList[6]), Double.Parse(temporarilyList[7]), temporarily, time, week, lectureRoom,
                       professorName, temporarilyList[8]));
                temporarilyList.Clear(); // 재사용을 위한 삭제
                #region 처음에 해본 방법
                /* if ((excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null) // 메뉴 번호에 따른 세부내용 확인
                     && (excelRange.Cells[i, 2] != null && excelRange.Cells[i, 2].Value2 != null)
                     && (excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].Value2 != null)
                     && (excelRange.Cells[i, 4] != null && excelRange.Cells[i, 4].Value2 != null)
                     && (excelRange.Cells[i, 5] != null && excelRange.Cells[i, 5].Value2 != null)
                     && (excelRange.Cells[i, 6] != null && excelRange.Cells[i, 6].Value2 != null)
                     && (excelRange.Cells[i, 7] != null && excelRange.Cells[i, 7].Value2 != null)
                     && (excelRange.Cells[i, 8] != null && excelRange.Cells[i, 8].Value2 != null)
                     && (excelRange.Cells[i, 9] != null && excelRange.Cells[i, 9].Value2 != null)
                     && (excelRange.Cells[i, 12] != null && excelRange.Cells[i, 12].Value2 != null))
                 {
                     nameTemporarily = excelRange.Cells[i, 5].Value2.ToString(); // private string lectureName; // 교과목명 5
                     temporarily = excelRange.Cells[i, 9].Value2.ToString(); //  private string date;// 전체시간 9 잠시 담아두기
                     if (excelRange.Cells[i, 10].Value2.ToString() != "") lectureRoom = excelRange.Cells[i, 10].Value2.ToString(); // private string lectureRoom;// 강의실 12 (엑셀기준 10)
                     else lectureRoom = "\t\t";

                     if (excelRange.Cells[i, 11].Value.ToString() != null) professorName = excelRange.Cells[i, 11].Value2.ToString(); // private string professorName; // 교수명 13(엑셀기준 11)

                     else professorName = "\t\t";

                     for (int k = 0; k < temporarily.Length; k++)
                     {
                         switch (temporarily[k]) //화목13:30-15:00,화18:00-20:00 이런거 나누기
                         {
                             case '월':
                             case '화':
                             case '수':
                             case '목':
                             case '금':
                                 week = week + temporarily[k];
                                 break;
                             default:
                                 time = time + temporarily[k];
                                 break;
                         }
                     }
                     //if (nameTemporarily.Length < 10)
                     //    nameTemporarily = nameTemporarily + "          ";

                     /*public LectureVO(int number, string department, string lectureNumber,
                     string lectureClassNumber, string lectureName, string completeDivision,
                     int grade, double credit, string[] dateTime, string[] lectureRoom,
                     string professorName, string lectureLanguage)

                     lecture.Add(new LectureVO(Int32.Parse(excelRange.Cells[i, 1].Value2.ToString()), excelRange.Cells[i, 2].Value2.ToString(), excelRange.Cells[i, 3].Value2.ToString(),
                      excelRange.Cells[i, 4].Value2.ToString(), nameTemporarily, excelRange.Cells[i, 6].Value2.ToString(),
                       Int32.Parse(excelRange.Cells[i, 7].Value2.ToString()), Double.Parse(excelRange.Cells[i, 8].Value2.ToString()), temporarily, time, week, lectureRoom,
                        professorName, excelRange.Cells[i, 12].Value2.ToString()));
                     week = "";
                     time = "";
                 }*/
                #endregion
            }
        }

        public bool SearchLectureWith() // 강의 출력을 위한 세부사항 설정
        {
            do// 어떤 세부 목록으로 검색할지 번호 선택
            {
                Console.Clear();
                Console.WriteLine("\n\n-----------------------------------------------------------------------------------------강의 출력-----------------------------------------------------------------------------------------");
                Console.WriteLine("\n\n\n");
                for (int i = 0; i < 10; i++)
                    Console.WriteLine("\t\t\t\t\t\t\t\t\t\t\t{0}. {1}", i, searchWithList[i]);
                Console.Write("\t\t\t\t\t\t\t\t\t\t\t어떤 것을 통해 보시겠습니까?  ");
                searchWithNumberString = Console.ReadLine();
            } while (!(searchWithNumberString.Equals("1") || searchWithNumberString.Equals("2") || searchWithNumberString.Equals("3") ||
              searchWithNumberString.Equals("4") || searchWithNumberString.Equals("5") || searchWithNumberString.Equals("6") ||
              searchWithNumberString.Equals("7") || searchWithNumberString.Equals("8") || searchWithNumberString.Equals("9") || searchWithNumberString.Equals("0")));// 어떤 세부 목록으로 검색할지 번호 선택
            if (searchWithNumberString.Equals("0")) return false;// 0입력시, 무한루프 나가기(뒤로가기)
            else if (searchWithNumberString.Equals("9"))// 전체
            {
                searchWithNumber = 0;
                SearchLecturePrint();
                return true; // (정상처리인지 확인용)
            }

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
                        Console.WriteLine("\t\t\t\t\t\t\t\t\t\t\t{0}. {1}", i, searchWithList[i]);
                    Console.WriteLine("\t\t\t\t\t\t\t\t\t\t\t어떤 것을 통해 보시겠습니까?  {0}", searchWithNumberString);
                }// 첫 do while이 아닌 경우(0~9이외의 값을 입력시) 메뉴를 다시 출력해야하므로
                Console.Write("\t\t\t\t\t\t\t\t\t\t\t{0}  ", searchWithList[Convert.ToInt32(searchWithNumberString)]);
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
                #endregion
            } while (check == 0);
            return true; // (정상처리인지 확인용)
        }
        public void SearchLecturePrint() // 강의 출력
        {
            for (int i = 1; i <= lecture.Count; i++)
            {
                if ((searchWithNumber != 0) && (excelRange.Cells[i, searchWithNumber].Value2.ToString() != searchWitinformation)) //searchWithNumber =0 이면 전체 출력이라 continue 하면 안됨
                    continue;

                check++; // 검색된 강의가 있는지 확인
                #region 강의출력 출력문
                if (lecture[i-1].Number != -1) // 각 행의 NO
                    Console.Write("{0,-5}", lecture[i-1].Number);
                if (lecture[i-1].Department != null) // 개설학과
                    Console.Write("{0,-10}\t", lecture[i-1].Department);
                if (lecture[i-1].LectureNumber != null) // 학수번호
                    Console.Write("{0,-10}", lecture[i-1].LectureNumber);
                if (lecture[i-1].LectureClassNumber != null) // 분반
                    Console.Write("{0,-5}", lecture[i-1].LectureClassNumber);
                if (lecture[i-1].LectureName != null)// 교과목명
                {
                    if (lecture[i-1].LectureName.Length > 20)
                        Console.Write("{0,-16}\t\t", lecture[i-1].LectureName);
                    else
                        Console.Write("{0,-14}\t\t\t", lecture[i-1].LectureName);
                }// 교과목명
                if (lecture[i-1].CompleteDivision != null) // 이수구분
                    Console.Write("{0,-8}", lecture[i-1].CompleteDivision);
                if (lecture[i-1].Grade != -1) // 학년
                    Console.Write("{0,-2}", lecture[i-1].Grade);
                if (lecture[i-1].Credit != -1) // 학점
                    Console.Write("{0,-2}", lecture[i-1].Credit);
                if (lecture[i-1].Date != null) //요일 및 강의시간
                {
                    if (lecture[i-1].Date.Length < 14)
                    {
                        Console.Write("{0,-15}\t\t", lecture[i-1].Date);
                    }
                    else
                    {
                        Console.Write("{0,-15}\t", lecture[i-1].Date);
                    }
                } //요일 및 강의시간
                if (lecture[i-1].LectureRoom != null) // 강의실
                    Console.Write("{0,-8}\t", lecture[i-1].LectureRoom);

                if (lecture[i-1].ProfessorName != null)
                { // 교수명
                    if (lecture[i-1].ProfessorName.Length < 1)
                    {
                        Console.Write("{0,-13}\t", lecture[i-1].ProfessorName);
                    }
                    else if (lecture[i-1].ProfessorName.Length < 3)
                    {
                        Console.Write("{0,-14}\t", lecture[i-1].ProfessorName);
                    }
                    else if (lecture[i-1].ProfessorName.Length > 4)
                    {
                        Console.Write("{0,-12}\t", lecture[i-1].ProfessorName);
                    }
                    else
                        Console.Write("{0,-16}", lecture[i-1].ProfessorName);
                }
                if (lecture[i-1].LectureLanguage != null) // 강의 언어
                    Console.Write("{0,-4}", lecture[i-1].LectureLanguage);
                #endregion
                //lecture[i-1].insertTimeSheet();

                Console.WriteLine("\r");
            }
            if (check != 0) // 검색된 강의가 있다면 0이 아니므로
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

                    if ((excelRange.Cells[i, 3].Value2.ToString.Equal(subjectNumber)) && (excelRange.Cells[i, 4].Value2.ToString.Equal(subjectclass)))
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
                else
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        Console.Write("{0,10}", selectLecture[j]);

                    }
                    Console.WriteLine(selectLecture);
                    Console.ReadLine();
                }
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
/*public void ReadExcelData(string path)
{ // path는 Excel파일의 전체 경로입니다.
  // 예. D:\test\test.xslx
    Excel.Application excelApp = null;
    Excel.Workbook wb = null;
    Excel.Worksheet ws = null;
    try
    {
        excelApp = new Excel.Application();
        wb = excelApp.Workbooks.Open(path);
        // path 대신 문자열도 가능합니다
        // 예. Open(@"D:\test\test.xslx");
        ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
        // 첫번째 Worksheet를 선택합니다.
        Excel.Range rng = ws.UsedRange;   // '여기'
        // 현재 Worksheet에서 사용된 셀 전체를 선택합니다.
        object[,] data = rng.Value;
        // 열들에 들어있는 Data를 배열 (One-based array)로 받아옵니다.
        for (int r = 1, i = 0; r <= data.GetLength(0); r++)
        {
            for (int c = 1; c <= data.GetLength(1); c++)
            {
                if (data[r, c] == null)
                {
                    continue;
                }
                // Data 빼오기
                // data[r, c] 는 excel의 (r, c) 셀 입니다.
                // data.GetLength(0)은 엑셀에서 사용되는 행의 수를 가져오는 것이고,
                // data.GetLength(1)은 엑셀에서 사용되는 열의 수를 가져오는 것입니다.
                // GetLength와 [ r, c] 의 순서를 바꿔서 사용할 수 있습니다.
            }
        }
        wb.Close(true);
        excelApp.Quit();
    }
    catch (Exception ex)
    {
        throw ex;
    }
    finally
    {
        ReleaseExcelObject(ws);
        ReleaseExcelObject(wb);
        ReleaseExcelObject(excelApp);
    }
}
*/
