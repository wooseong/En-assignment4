using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace schedule
{
    class LectureVO
    {
        private int number; //NO
        private string department; //개설학과전공 
        private string lectureNumber; // 학수번호
        private string lectureClassNumber; // 분반
        private string lectureName; // 교과목명
        private string completeDivision; // 이수구분
        private int grade; // 학년
        private double credit; // 학점
        private string[] dateTime;// = new string[3];// 요일 및 시간
        private string[] lectureRoom;// = new string[2]; // 강의실
        private string professorName; // 교수명
        private string lectureLanguage;//강의 언어

        public LectureVO(int number, string department, string lectureNumber,
            string lectureClassNumber, string lectureName, string completeDivision,
            int grade, double credit, string[] dateTime, string[] lectureRoom,
            string professorName, string lectureLanguage)
        {
            this.number = number;
            this.department = department;
            this.lectureNumber = lectureNumber;
            this.lectureClassNumber = lectureClassNumber;
            this.lectureName = lectureName;
            this.completeDivision = completeDivision;
            this.grade = grade;
            this.credit = credit;
            this.dateTime = dateTime;
            this.lectureRoom = lectureRoom;
            this.professorName = professorName;
            this.lectureLanguage = lectureLanguage;

        } // 생성자

        public int Number
        {
            get { return number; }
            set { number = value; }
        }
        public string Department
        {
            get { return department; }
            set { department = value; }
        }
        public string LectureNumber
        {
            get { return lectureNumber; }
            set { LectureNumber = value; }
        }
        public string LectureClassNumber
        {
            get { return lectureClassNumber; }
            set { lectureClassNumber = value; }
        }
        public string LectureName
        {
            get { return lectureName; }
            set { lectureName = value; }
        }
        public string CompleteDivision
        {
            get { return completeDivision; }
            set { completeDivision = value; }
        }
        public int Grade
        {
            get { return grade; }
            set { grade = value; }
        }
        public double Credit
        {
            get { return credit; }
            set { credit = value; }
        }
        public string[] DateTime
        {
            get { return dateTime; }
            set { dateTime = value; }
        }
        public string[] LectureRoom
        {
            get { return lectureRoom; }
            set { lectureRoom = value; }
        }
        public string ProfessorName
        {
            get { return professorName; }
            set { professorName = value; }
        }
        public string LectureLanguage
        {
            get { return lectureLanguage; }
            set { lectureLanguage = value; }
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
