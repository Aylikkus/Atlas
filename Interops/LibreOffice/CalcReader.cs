using Atlas.Data;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Documents;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.sheet;
using unoidl.com.sun.star.table;
using unoidl.com.sun.star.text;
using unoidl.com.sun.star.uno;

namespace Atlas.Interops.LibreOffice
{
    public class CalcReader : IDocReader, IDisposable
    {
        XComponentContext xContext = Bootstrap.bootstrap();
        XComponent xComp;

        public void Dispose()
        {
            xComp.dispose();
        }

        void parseLessonCell(string cellCnt, int sem, ref WorkInfo les, in SemesterInfo si)
        {
            bool parsed = int.TryParse(cellCnt, out int hours);

            if (parsed)
            {
                if (les == null)
                    les = new WorkInfo(si);

                les.SetOn(sem, hours);
            }
        }

        Dictionary<string, string> parseCompetentionCell(string cellCnt,
            in Dictionary<string, string> allComps)
        {
            string[] comps = cellCnt?.Split(new char[] { ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            Dictionary<string, string> discComps = new Dictionary<string, string>();
            foreach (var c in comps)
            {
                discComps[c] = allComps[c];
            }

            return discComps;
        }

        private void loadLibreOffice(string pathToFile)
        {
            XMultiComponentFactory xMCF = xContext.getServiceManager();
            object oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
            XComponentLoader xLoader = (XComponentLoader)oDesktop;
            PropertyValue[] args = new PropertyValue[2];
            args[0] = new PropertyValue();
            args[0].Name = "ReadOnly";
            args[0].Value.setValue(typeof(bool), true);

            args[1] = new PropertyValue();
            args[1].Name = "Hidden";
            args[1].Value.setValue(typeof(bool), true);

            xComp = xLoader.loadComponentFromURL(@"file:///" + pathToFile, "_blank", 0, args);
        }

        private XSpreadsheet getByIndex(int nIndex)
        {
            XSpreadsheetDocument doc = (XSpreadsheetDocument)xComp;
            XSpreadsheets sheets = doc.getSheets();
            XIndexAccess indexAccess = (XIndexAccess)sheets;
            return (XSpreadsheet)indexAccess.getByIndex(nIndex).Value;
        }

        public DocAttributes PullAttributes(string pathToFile)
        {
            loadLibreOffice(pathToFile);

            // Нумерация листов с нуля, ячеек тоже
            XSpreadsheet titul = getByIndex(0);
            string departament = titul.getCellByPosition(2, 26).getFormula();
            string faculty = titul.getCellByPosition(2, 27).getFormula();
            string specialization = Formatter.GetSpecialization(
                titul.getCellByPosition(2, 18).getFormula());
            string profile = Formatter.GetProfile(
                titul.getCellByPosition(2, 19).getFormula());

            string grLevel = titul.getCellByPosition(1, 29).getFormula();
            string edLevel = Formatter.GetEducationLevel(grLevel);
            string edType = Formatter.GetEducationType(
                titul.getCellByPosition(1, 31).getFormula());
            string tmp = titul.getCellByPosition(20, 29).getFormula();
            int year = int.Parse(titul.getCellByPosition(20, 29).getFormula()
                .Trim(new char[] { '\'', ' ', '\"'}));

            // Компетенции
            XSpreadsheet comps = getByIndex(4);

            Dictionary<string, string> allComps = new Dictionary<string, string>(128);
            for (int compRow = 1; compRow < 1000; compRow++)
            {
                if (comps.getCellByPosition(1, compRow).getFormula() != "")
                {
                    string compName = comps.getCellByPosition(1, compRow).getFormula().Trim();
                    allComps[compName] = comps.getCellByPosition(3, compRow).getFormula();
                }
            }

            // План
            XSpreadsheet plan = getByIndex(3);

            List<Discipline> disciplines = new List<Discipline>(128);
            for (int discRow = 1; discRow < 1000; discRow++)
            {
                string discCode = plan.getCellByPosition(1, discRow).getFormula();
                string discName = plan.getCellByPosition(2, discRow).getFormula();

                if (discCode == "" || discName == "")
                    continue;

                // Проверка на группировку дисциплин
                if (discCode[0] == 'Б' && !discName.ToLowerInvariant().Contains("дисциплины"))
                {
                    SemesterInfo si = new SemesterInfo();

                    // Экз, Зач, ЗачОц, КурПр, КурРаб, РГР
                    WorkInfo[] examInfos = new WorkInfo[6];

                    for (int j = 3; j < 9; j++)
                    {
                        string workInfoSems = plan.getCellByPosition(j, discRow).getFormula();
                        if (workInfoSems == null) workInfoSems = "";
                        workInfoSems = workInfoSems.Trim(new char[] { '\'', ' ', '\"' });

                        if (workInfoSems.Length > 0)
                            examInfos[j - 3] = new WorkInfo(si);

                        foreach (char c in workInfoSems)
                        {
                            int n = int.Parse(c.ToString(), NumberStyles.HexNumber);
                            examInfos[j - 3].SetOn(n, 0);
                        }
                    }

                    // Лекции, Лабы, Практики, Сам.Работы, Контроль
                    WorkInfo[] lesInfos = new WorkInfo[5];

                    
                    for (int j = 17, s = 1; 
                        plan.getCellByPosition(j, 2).getFormula() == "з.е." && s <= 16; 
                        j += 7, s++)
                    {
                        parseLessonCell(plan.getCellByPosition(j + 2, 2).getFormula(), s, ref lesInfos[0], si);
                        parseLessonCell(plan.getCellByPosition(j + 3, 2).getFormula(), s, ref lesInfos[1], si);
                        parseLessonCell(plan.getCellByPosition(j + 4, 2).getFormula(), s, ref lesInfos[2], si);
                        parseLessonCell(plan.getCellByPosition(j + 5, 2).getFormula(), s, ref lesInfos[3], si);
                        parseLessonCell(plan.getCellByPosition(j + 6, 2).getFormula(), s, ref lesInfos[4], si);
                    }

                    
                    Dictionary<string, string> tmpComp = parseCompetentionCell(
                        plan.getCellByPosition(103, discRow).getFormula(), allComps);

                    Discipline disc = new Discipline(discCode, discName);
                    disc.Semester = si;
                    disc.Exam = examInfos[0];
                    disc.Credits = examInfos[1];
                    disc.RatedCredits = examInfos[2];
                    disc.CourseProjects = examInfos[3];
                    disc.CourseWorks = examInfos[4];
                    disc.RGR = examInfos[5];
                    disc.Lectures = lesInfos[0];
                    disc.Laboratory = lesInfos[1];
                    disc.Practice = lesInfos[2];
                    disc.Independent = lesInfos[3];
                    disc.Control = lesInfos[4];
                    disc.Competentions = tmpComp;
                    disciplines.Add(disc);
                }
            }

            DocAttributes da = new DocAttributes(specialization, profile);
            da.Departament = departament;
            da.Faculty = faculty;
            da.EducationLevel = edLevel;
            da.GraduationLevel = grLevel;
            da.EducationType = edType;
            da.YearOfEntrance = year;
            da.Competentions = allComps;
            da.Disciplines = disciplines;

            int x = 0;

            return da;
        }
    }
}
