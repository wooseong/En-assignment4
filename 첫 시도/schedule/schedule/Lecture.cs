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
        double maxCredit = 21;
        double maxIntrestCredit = 24;

        private string searchWithNumberString; // SearchLecturePrintWith의 메뉴 번호 선택 변수
        private int searchWithNumber; // SearchLecturePrintWith의 메뉴 번호에 따른 엑셀 열 번호 단, 0이면 전체 출력
        private string searchWitinformation; //  SearchLecturePrintWith의 새부 검색 내용 변수
        private int check; //새부 검색 내용이 출력 되었는지 확인
        List<LectureVO> lecture = new List<LectureVO>(); // 엑셀로 부터 강의 모두 받아온 list

        private string subjectNumber;// 수강 신청 학수번호
        private string subjectclass;// 수강 신청 분반

        private string mainMenuNumber;//출력 후 2: 수강신청, 5: 관심과목 담기, 0: 뒤로
        private int time = 9;
        List<List<string>> sheet = new List<List<string>>();

        List<LectureVO> selectLecture = new List<LectureVO>(); // 수강 신청한 내 강의목록
        List<LectureVO> selectIntrestLecture = new List<LectureVO>(); // 관심과목 담은 내 목록
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
            }
            else
            {
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
            }
            //#region
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t\t 2.수강 신청"); // 엑셀에 추가
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t\t 5.관심 과목 추가");
            Console.WriteLine("\t\t\t\t\t\t\t\t\t\t\t 그 외. 뒤로가기");
            mainMenuNumber = Console.ReadLine();
            if (mainMenuNumber.Equals("2"))
            {
                AddLecture(22);
            }
            else if (mainMenuNumber.Equals("5"))
            {
                AddLecture(55);
            }
            return true; // (정상처리인지 확인용)
        }


        public void SearchLecturePrint() // 강의 출력
        {
            for (int i = 2; i <= lecture.Count; i++)
            {
                if ((searchWithNumber != 0) && (excelRange.Cells[i, searchWithNumber].Value2.ToString() != searchWitinformation)) //searchWithNumber =0 이면 전체 출력이라 continue 하면 안됨
                    continue;

                check++; // 검색된 강의가 있는지 확인
                #region 강의출력 출력문
                if (lecture[i - 2].Number != -1) // 각 행의 NO
                    Console.Write("{0,-5}", lecture[i - 2].Number);
                if (lecture[i - 2].Department != null) // 개설학과
                    Console.Write("{0,-10}\t", lecture[i - 2].Department);
                if (lecture[i - 2].LectureNumber != null) // 학수번호
                    Console.Write("{0,-10}", lecture[i - 2].LectureNumber);
                if (lecture[i - 2].LectureClassNumber != null) // 분반
                    Console.Write("{0,-5}", lecture[i - 2].LectureClassNumber);
                if (lecture[i - 2].LectureName != null)// 교과목명
                {
                    if (lecture[i - 2].LectureName.Length > 20)
                        Console.Write("{0,-16}\t\t", lecture[i - 2].LectureName);
                    else
                        Console.Write("{0,-14}\t\t\t", lecture[i - 2].LectureName);
                }// 교과목명
                if (lecture[i - 2].CompleteDivision != null) // 이수구분
                    Console.Write("{0,-8}", lecture[i - 2].CompleteDivision);
                if (lecture[i - 2].Grade != -1) // 학년
                    Console.Write("{0,-2}", lecture[i - 2].Grade);
                if (lecture[i - 2].Credit != -1) // 학점
                    Console.Write("{0,-2}", lecture[i - 2].Credit);
                if (lecture[i - 2].Date != null) //요일 및 강의시간
                {
                    if (lecture[i - 2].Date.Length < 14)
                    {
                        Console.Write("{0,-15}\t\t", lecture[i - 2].Date);
                    }
                    else
                    {
                        Console.Write("{0,-15}\t", lecture[i - 2].Date);
                    }
                } //요일 및 강의시간
                if (lecture[i - 2].LectureRoom != null) // 강의실
                    Console.Write("{0,-8}\t", lecture[i - 2].LectureRoom);
                if (lecture[i - 2].ProfessorName != null)
                { // 교수명
                    if (lecture[i - 2].ProfessorName.Length < 1)
                    {
                        Console.Write("{0,-13}\t", lecture[i - 2].ProfessorName);
                    }
                    else if (lecture[i - 2].ProfessorName.Length < 3)
                    {
                        Console.Write("{0,-14}\t", lecture[i - 2].ProfessorName);
                    }
                    else if (lecture[i - 2].ProfessorName.Length > 4)
                    {
                        Console.Write("{0,-12}\t", lecture[i - 2].ProfessorName);
                    }
                    else
                        Console.Write("{0,-16}", lecture[i - 2].ProfessorName);
                }
                if (lecture[i - 2].LectureLanguage != null) // 강의 언어
                    Console.Write("{0,-4}", lecture[i - 2].LectureLanguage);
                #endregion

                Console.WriteLine("\r");
            }
        }
        public void AddLecture(int menuNumber) // menuNumber 2,22 : 수강신청, 5,55 : 관심과목 추가
        {
            do
            {
                int ox = 0, ox1 = 0;
                if (menuNumber != 22 && menuNumber != 55) Console.Clear();
                Console.Write("\t수강 신청하실 과목의 학수번호와 분반 입력하세요.  \n\t학수 번호  :  ");
                subjectNumber = Console.ReadLine();
                if (subjectNumber.Equals("0")) break;
                Console.Write("\t분     반  :  ");
                subjectclass = Console.ReadLine();
                if (subjectclass.Equals("0")) break;
                check = 0;
                for (int i = 2; i <= lecture.Count; i++)
                {
                    //excelRange.Cells[i, searchWithNumber].Value2.ToString()
                    if (!(excelRange.Cells[i, 3].Value2.ToString() != subjectNumber) && !(excelRange.Cells[i, 4].Value2.ToString() != subjectclass))
                    {
                        check++;
                        if (menuNumber == 2 || menuNumber == 22) // 수강신청
                        {
                            for (int j = 0; j < selectLecture.Count; j++)
                            {
                                if (!(selectLecture[j].LectureNumber != subjectNumber))
                                {
                                    ox = 1;
                                    Console.Write("같은 과목이 이미 존재합니다.");
                                    break;
                                }
/*                                else
                                {
                                    selectIntrestLecture[j].DateTime
                                }*/
                            }
                            if (ox == 0)
                            {
                                maxCredit -= lecture[i - 2].Credit;
                                selectLecture.Add(new LectureVO(lecture[i - 2].Number, lecture[i - 2].Department, lecture[i - 2].LectureNumber,
                                    lecture[i - 2].LectureClassNumber, lecture[i - 2].LectureName, lecture[i - 2].CompleteDivision,
                                    lecture[i - 2].Grade, lecture[i - 2].Credit, lecture[i - 2].Date, lecture[i - 2].DateTime, lecture[i - 2].DateWeek,
                                    lecture[i - 2].LectureRoom, lecture[i - 2].ProfessorName, lecture[i - 2].LectureLanguage));
                                Console.WriteLine("정상적으로 추가되셨습니다.");
                                Console.Write("수강 가능한 남은 학점은 {0}학점", maxCredit);
                            }
                        }
                        else if (menuNumber == 5 || menuNumber == 55)//관심과목담기
                        {
                            for (int j = 0; j < selectIntrestLecture.Count; j++)
                            {
                                if (!(selectIntrestLecture[j].LectureNumber != subjectNumber))
                                {
                                    ox = 1;
                                    Console.Write("같은 과목이 이미 존재합니다.");
                                }
                            }
                            if (ox == 0)
                            {
                                maxIntrestCredit -= lecture[i - 2].Credit;
                                selectIntrestLecture.Add(new LectureVO(lecture[i - 2].Number, lecture[i - 2].Department, lecture[i - 2].LectureNumber,
                                    lecture[i - 2].LectureClassNumber, lecture[i - 2].LectureName, lecture[i - 2].CompleteDivision,
                                    lecture[i - 2].Grade, lecture[i - 2].Credit, lecture[i - 2].Date, lecture[i - 2].DateTime, lecture[i - 2].DateWeek,
                                    lecture[i - 2].LectureRoom, lecture[i - 2].ProfessorName, lecture[i - 2].LectureLanguage));
                                Console.Write("정상적으로 추가되셨습니다.");
                                Console.Write("수강 가능한 남은 학점은 {0}학점", maxIntrestCredit);
                            }
                        }
                        Console.ReadLine();
                        break;
                    }
                }
                if (check == 0)
                {
                    Console.Write("그런 수업은 없습니다.");
                    Thread.Sleep(1000);
                }
            } while (check == 0);
        }
        public void ErasureLecture(int menuNumber)
        {
            PrintSelectLectureList(menuNumber);
            if (!(menuNumber == 3 && maxCredit == 21) || (menuNumber == 6 && maxIntrestCredit == 24))
            {
                do
                {
                    //Console.Clear();
                    Console.Write("\t0: 뒤로가기\n\t삭제하실 과목의 학수번호와 분반 입력하세요.  \n\t학수 번호  :  ");
                    subjectNumber = Console.ReadLine();
                    if (subjectNumber.Equals("0")) break;
                    Console.Write("\t분     반  :  ");
                    subjectclass = Console.ReadLine();
                    if (subjectclass.Equals("0")) break;
                    check = 1;
                    if (menuNumber == 3)
                    {
                        for (int i = 0; i <= selectLecture.Count; i++)
                        {
                            if (!(selectLecture[i].LectureNumber != subjectNumber) && (selectLecture[i].LectureNumber != subjectclass))
                            {
                                check--;
                                maxCredit += selectLecture[i].Credit;
                                selectLecture.RemoveAt(i);
                                break;
                            }
                        }
                        Console.WriteLine("정상적으로 삭제되셨습니다.");
                        Console.Write("수강 가능한 남은 학점은 {0}학점", maxCredit);
                    }
                    else if (menuNumber == 6)
                    {
                        for (int i = 0; i <= selectIntrestLecture.Count; i++)
                        {
                            if (!(selectIntrestLecture[i].LectureNumber != subjectNumber) && (selectIntrestLecture[i].LectureNumber != subjectclass))
                            {
                                check--;
                                maxIntrestCredit += selectIntrestLecture[i].Credit;
                                selectIntrestLecture.RemoveAt(i);
                                break;
                            }
                        }
                        Console.WriteLine("정상적으로 삭제되셨습니다.");
                        Console.Write("수강 가능한 남은 학점은 {0}학점", maxCredit);
                    }
                    Console.ReadLine();
                    if (check == 1)
                    {
                        Console.Write("그런 수업은 없습니다.");
                        Thread.Sleep(1000);
                    }
                    else break;
                } while (check == 1);
            }
        }
        public void PrintSelectLectureList(int menuNumber)
        {
            Console.Clear();
            if ( menuNumber == 3)
            {
                if (maxCredit == 21)
                {
                    Console.Write("등록하신 과목이 없습니다.");
                    Console.ReadLine();
                }
                else
                {
                    for (int i = 0; i < selectLecture.Count; i++)
                    {
                        #region 강의출력 출력문
                        if (selectLecture[i].Number != -1) // 각 행의 NO
                            Console.Write("{0,-5}", selectLecture[i].Number);
                        if (selectLecture[i].Department != null) // 개설학과
                            Console.Write("{0,-10}\t", selectLecture[i].Department);
                        if (selectLecture[i].LectureNumber != null) // 학수번호
                            Console.Write("{0,-10}", selectLecture[i].LectureNumber);
                        if (selectLecture[i].LectureClassNumber != null) // 분반
                            Console.Write("{0,-5}", selectLecture[i].LectureClassNumber);
                        if (selectLecture[i].LectureName != null)// 교과목명
                        {
                            if (selectLecture[i].LectureName.Length > 20)
                                Console.Write("{0,-16}\t\t", selectLecture[i].LectureName);
                            else
                                Console.Write("{0,-14}\t\t\t", selectLecture[i].LectureName);
                        }// 교과목명
                        if (selectLecture[i].CompleteDivision != null) // 이수구분
                            Console.Write("{0,-8}", selectLecture[i].CompleteDivision);
                        if (selectLecture[i].Grade != -1) // 학년
                            Console.Write("{0,-2}", selectLecture[i].Grade);
                        if (selectLecture[i].Credit != -1) // 학점
                            Console.Write("{0,-2}", selectLecture[i].Credit);
                        if (selectLecture[i].Date != null) //요일 및 강의시간
                        {
                            if (selectLecture[i].Date.Length < 14)
                            {
                                Console.Write("{0,-15}\t\t", selectLecture[i].Date);
                            }
                            else
                            {
                                Console.Write("{0,-15}\t", selectLecture[i].Date);
                            }
                        } //요일 및 강의시간
                        if (selectLecture[i].LectureRoom != null) // 강의실
                            Console.Write("{0,-8}\t", selectLecture[i].LectureRoom);
                        if (selectLecture[i].ProfessorName != null)
                        { // 교수명
                            if (selectLecture[i].ProfessorName.Length < 1)
                            {
                                Console.Write("{0,-13}\t", selectLecture[i].ProfessorName);
                            }
                            else if (selectLecture[i].ProfessorName.Length < 3)
                            {
                                Console.Write("{0,-14}\t", selectLecture[i].ProfessorName);
                            }
                            else if (selectLecture[i].ProfessorName.Length > 4)
                            {
                                Console.Write("{0,-12}\t", selectLecture[i].ProfessorName);
                            }
                            else
                                Console.Write("{0,-16}", selectLecture[i].ProfessorName);
                        }
                        if (selectLecture[i].LectureLanguage != null) // 강의 언어
                            Console.Write("{0,-4}", selectLecture[i].LectureLanguage);
                        #endregion

                        Console.WriteLine("\r");
                    }
                }
            }
            else if (menuNumber == 4 || menuNumber == 6)
            {

                if (maxIntrestCredit == 24)
                {
                    Console.Write("등록하신 과목이 없습니다.");
                    Console.ReadLine();
                }
                else
                {
                    for (int i = 0; i < selectIntrestLecture.Count; i++)
                    {
                        #region 강의출력 출력문
                        if (selectIntrestLecture[i].Number != -1) // 각 행의 NO
                            Console.Write("{0,-5}", selectIntrestLecture[i].Number);
                        if (selectIntrestLecture[i].Department != null) // 개설학과
                            Console.Write("{0,-10}\t", selectIntrestLecture[i].Department);
                        if (selectIntrestLecture[i].LectureNumber != null) // 학수번호
                            Console.Write("{0,-10}", selectIntrestLecture[i].LectureNumber);
                        if (selectIntrestLecture[i].LectureClassNumber != null) // 분반
                            Console.Write("{0,-5}", selectIntrestLecture[i].LectureClassNumber);
                        if (selectIntrestLecture[i].LectureName != null)// 교과목명
                        {
                            if (selectIntrestLecture[i].LectureName.Length > 20)
                                Console.Write("{0,-16}\t\t", selectIntrestLecture[i].LectureName);
                            else
                                Console.Write("{0,-14}\t\t\t", selectIntrestLecture[i].LectureName);
                        }// 교과목명
                        if (selectIntrestLecture[i].CompleteDivision != null) // 이수구분
                            Console.Write("{0,-8}", selectIntrestLecture[i].CompleteDivision);
                        if (selectIntrestLecture[i].Grade != -1) // 학년
                            Console.Write("{0,-2}", selectIntrestLecture[i].Grade);
                        if (selectIntrestLecture[i].Credit != -1) // 학점
                            Console.Write("{0,-2}", selectIntrestLecture[i].Credit);
                        if (selectIntrestLecture[i].Date != null) //요일 및 강의시간
                        {
                            if (selectIntrestLecture[i].Date.Length < 14)
                            {
                                Console.Write("{0,-15}\t\t", selectIntrestLecture[i].Date);
                            }
                            else
                            {
                                Console.Write("{0,-15}\t", selectIntrestLecture[i].Date);
                            }
                        } //요일 및 강의시간
                        if (selectIntrestLecture[i].LectureRoom != null) // 강의실
                            Console.Write("{0,-8}\t", selectIntrestLecture[i].LectureRoom);
                        if (selectIntrestLecture[i].ProfessorName != null)
                        { // 교수명
                            if (selectIntrestLecture[i].ProfessorName.Length < 1)
                            {
                                Console.Write("{0,-13}\t", selectIntrestLecture[i].ProfessorName);
                            }
                            else if (selectIntrestLecture[i].ProfessorName.Length < 3)
                            {
                                Console.Write("{0,-14}\t", selectIntrestLecture[i].ProfessorName);
                            }
                            else if (selectIntrestLecture[i].ProfessorName.Length > 4)
                            {
                                Console.Write("{0,-12}\t", selectIntrestLecture[i].ProfessorName);
                            }
                            else
                                Console.Write("{0,-16}", selectIntrestLecture[i].ProfessorName);
                        }
                        if (selectIntrestLecture[i].LectureLanguage != null) // 강의 언어
                            Console.Write("{0,-4}", selectIntrestLecture[i].LectureLanguage);
                        #endregion

                        Console.WriteLine("\r");
                    }
                }
            }
        }// 관심과목 출력

        public void initTimeSheet()
        {
            string[] week = { "월      \t", "화      \t", "수      \t", "목      \t", "금      \t" };
            for (int i = 0; i < 24; i++)
            {
                sheet.Add(new List<string>());

                for (int j = 0; j < 6; j++)
                {
                    if (i == 0 && j != 0)
                    {
                        sheet[i].Add(week[j - 1]);
                    }
                    else if (i != 0 && j == 0)
                    {
                        if (i % 2 == 1) sheet[i].Add(time + ":00-" + time + ":30   \t");
                        else
                        {
                            sheet[i].Add(time + ":30-" + time + ":00  \t");
                            time++;
                        }
                    }
                    else
                        sheet[i].Add("        \t");
                }
            }
        }
        public void printTimeSheet()
        {

            string time = "";


            for (int i = 0; i < selectLecture.Count; i++)
            {
                if (selectLecture[i].DateTime.Length < 12)
                {
                    time = selectLecture[i].DateTime.Substring(0, 5);
                    for (int j = 0; j < selectLecture[i].DateWeek.Length; j++)
                    {
                        switch (selectLecture[i].DateWeek[j])
                        {
                            case '월':
                                TimeCase(i, time, 1);
                                break;

                            case '화':
                                TimeCase(i, time, 2);
                                break;
                            case '수':
                                TimeCase(i, time, 3);
                                break;
                            case '목':
                                TimeCase(i, time, 4);
                                break;
                            case '금':
                                TimeCase(i, time, 5);
                                break;
                        }
                    }
                }
            }
            for (int i = 0; i < sheet.Count; i++)
            {
                for (int j = 0; j < sheet[i].Count; j++)
                {
                    Console.Write(sheet[i][j] + "\t");
                }
                Console.WriteLine();
            }
        }

        public void TimeCase(int i, string time, int selWeek)
        {
            Console.WriteLine(time);
            switch (time)
            {
                case "09:00":
                    sheet[1][selWeek] = selectIntrestLecture[i].LectureName + " " + selectIntrestLecture[i].LectureRoom;
                    break;
                case "10:00":
                    sheet[3][selWeek] = selectIntrestLecture[i].LectureName + " " + selectIntrestLecture[i].LectureRoom;
                    break;
                case "10:30":
                    sheet[4][selWeek] = selectIntrestLecture[i].LectureName + " " + selectIntrestLecture[i].LectureRoom;
                    break;
                case "12:00":
                    sheet[7][selWeek] = selectIntrestLecture[i].LectureName + " " + selectIntrestLecture[i].LectureRoom;
                    break;
                case "13:30":
                    sheet[10][selWeek] = selectIntrestLecture[i].LectureName + " " + selectIntrestLecture[i].LectureRoom;
                    break;
                case "14:00":
                    sheet[11][selWeek] = selectIntrestLecture[i].LectureName + " " + selectIntrestLecture[i].LectureRoom;
                    break;
                case "15:00":
                    sheet[13][selWeek] = selectIntrestLecture[i].LectureName + " " + selectIntrestLecture[i].LectureRoom;
                    break;
                case "16:00":
                    sheet[15][selWeek] = selectIntrestLecture[i].LectureName + " " + selectIntrestLecture[i].LectureRoom;
                    break;
                case "18:00":
                    sheet[19][selWeek] = selectIntrestLecture[i].LectureName + " " + selectIntrestLecture[i].LectureRoom;
                    break;

            }
        }
    }
}