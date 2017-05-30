using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace schedule
{
    class TimeSheet
    {
        private int time = 9;
        
        List<List<string>> sheet = new List<List<string>>();
        public TimeSheet()
        {
           
            string[] week = { "월      \t", "화      \t", "수      \t", "목      \t", "금      \t" };
            for (int i = 0; i < 24; i++)
            {
                sheet.Add(new List<string>()); 

                for (int j = 0; j < 6; j++)
                {
                    if(i == 0 && j != 0)
                    {
                        sheet[i].Add(week[j-1]);
                    }
                    else if (i != 0 && j == 0)
                    {
                        if (i % 2 == 1) sheet[i].Add(time + "시     \t");
                        else
                        {
                            sheet[i].Add(time + "시 반  \t");
                            time++;
                        }
                    }
                    else
                    sheet[i].Add(".       \t"); 
                }
            }
        }

        public void printTimeSheet()
        {
            for(int i = 0; i < sheet.Count; i++)
            {
                for(int j = 0; j < sheet[i].Count; j++)
                {
                    Console.Write(sheet[i][j] + "\t");
                }
                Console.WriteLine();
            }
           

        }
        public void AddTimeSheet(List<string> selected)
        {
            int[] x = new int[2];
               int y;
            int j = 0;
            for(int i = 0; i < selected[9].Length; i++)
            {
                switch (selected[9][i].ToString())
                {
                    case "월":
                        x[j++] = 2;
                        break;

                    case "화":
                        x[j++] = 3;
                        break;

                    case "수":
                        x[j++] = 4;
                        break;

                    case "목":
                        x[j++] = 5;
                        break;

                    case "금":
                        x[j++] = 6;
                        break;

                }
                
            }

        }
    }
   
}
