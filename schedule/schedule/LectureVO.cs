using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace schedule
{
    class LectureVO
    {
        private int number = -1; //NO 1
        private string department; //개설학과전공 2
        private string lectureNumber; // 학수번호 3
        private string lectureClassNumber; // 분반 4
        private string lectureName; // 교과목명 5
        private string completeDivision; // 이수구분 6
        private int grade = -1; // 학년 7
        private double credit = -1; // 학점 8 
        private string date;// 전체시간 9
        private string dateTime;// 시간부분 10
        private string dateWeek; //요일부분 11
        private string lectureRoom;// 강의실 12
        private string professorName; // 교수명 13
        private string lectureLanguage;//강의 언어 14

        public LectureVO()
        {
            this.number = -1;
            this.department = null;
            this.lectureNumber = null;
            this.lectureClassNumber = null;
            this.lectureName = null;
            this.completeDivision = null;
            this.grade = -1;
            this.credit = -1.0;
            this.date = null;
            this.dateTime = null;
            this.dateWeek = null;
            this.lectureRoom = null;
            this.professorName = null;
            this.lectureLanguage = null;

        }
        public LectureVO(int number, string department, string lectureNumber,
            string lectureClassNumber, string lectureName, string completeDivision,
            int grade, double credit, string date, string dateTime, string dateWeek, string lectureRoom,
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
            this.date = date;
            this.dateTime = dateTime;
            this.dateWeek = dateWeek;
            this.lectureRoom = lectureRoom;
            this.professorName = professorName;
            this.lectureLanguage = lectureLanguage;

        } // 생성자

        public void insertTimeSheet()
        {
            int i = 0;
            string stemporarily;
            List<string> temporarily = new List<string>();
            while (i < dateTime.Length)
            {
                if (dateTime[i] != '-' || dateTime[i] != ' ' || dateTime[i] != ',')
                {
                    stemporarily = dateTime.Substring(i, 1);
                    Console.WriteLine(stemporarily);
                    temporarily.Add(DateTime.Substring(i, 1));
                    i = i + 1;
                }
                 else
                 {
                     i++;
                 }
            }

            for(int j = 0; j < temporarily.Count; j++)
             {
                 Console.WriteLine(temporarily[j]);
             }

        }
        public void inserFavorit(LectureVO lecture) // 생성자가 
        {
            this.number = lecture.Number;
            this.department = lecture.Department;
            this.lectureNumber = lecture.LectureNumber;
            this.lectureClassNumber = lecture.LectureClassNumber;
            this.lectureName = lecture.LectureName;
            this.completeDivision = lecture.CompleteDivision;
            this.grade = lecture.Grade;
            this.credit = lecture.Credit;
            this.date = lecture.Date;
            this.dateTime = lecture.DateTime;
            this.dateWeek = lecture.DateWeek;
            this.lectureRoom = lecture.LectureRoom;
            this.professorName = lecture.ProfessorName;
            this.lectureLanguage = lecture.LectureLanguage;
        }
    #region
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
        public string Date
        {
            get { return date; }
            set { date = value; }
        }
        public string DateTime
        {
            get { return dateTime; }
            set { dateTime = value; }

        }
        public string DateWeek
        {
            get { return dateWeek; }
            set { dateWeek = value; }
        }
        public string LectureRoom
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
    #endregion
    }



}