using Atlas.Data;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.frame.status;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.text;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.util;
using unoidl.com.sun.star.xml.dom;
using unoidl.com.sun.star.xml.sax;

namespace Atlas.Interops.LibreOffice
{
    public class WriterGenerator : IDocGenerator, IDisposable
    {
        XComponentContext xContext = Bootstrap.bootstrap();
        XComponent xComp;

        string formatSemArray(int[] arr)
        {
            StringBuilder semBld = new StringBuilder();
            if (arr.Length > 1)
            {
                for (int i = 0; i < arr.Length; i++)
                {
                    semBld.Append(arr[i] + (i == arr.Length - 1 ? "" : ", "));
                }
            }
            else
            {
                semBld.Append(arr[0]);
            }

            return semBld.ToString();
        }

        string formatAttestation(WorkInfo exam, WorkInfo credits, WorkInfo ratedCredits)
        {
            List<string> strs = new List<string>(3);

            if (credits != null) strs.Add("зачёт");
            if (ratedCredits != null) strs.Add("зачёт с оценкой");
            if (exam != null) strs.Add("экзамен");

            string att = string.Join(", ", strs.ToArray());
            return char.ToUpper(att[0]) + att.Substring(1);
        }

        int getTotalDisc(in Discipline disc)
        {
            WorkInfo[] works = new WorkInfo[]
            {
                disc.Lectures,
                disc.Practice,
                disc.Laboratory,
                disc.Independent,
                disc.Control,
            };

            int count = 0;

            foreach (var w in works)
            {
                if (w != null)
                    count += w.Total;
            }

            return count;
        }

        bool findReplace(string text, string repl)
        {
            XReplaceable xRepl = xComp as XReplaceable;
            XReplaceDescriptor replDesc = xRepl.createReplaceDescriptor();
            replDesc.setSearchString(text);
            replDesc.setReplaceString(repl);

            return xRepl.replaceAll(replDesc) == 1;
        }

        void loadWriter(string pathToFile)
        {
            XMultiComponentFactory xMCF = xContext.getServiceManager();
            object oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
            XComponentLoader xLoader = (XComponentLoader)oDesktop;
            PropertyValue[] args = new PropertyValue[2];
            args[0] = new PropertyValue();
            args[0].Name = "Hidden";
            args[0].Value.setValue(typeof(bool), true);

            args[1] = new PropertyValue();
            args[1].Name = "ReadOnly";
            args[1].Value.setValue(typeof(bool), true);

            xComp = xLoader.loadComponentFromURL(@"file:///" + pathToFile, "_blank", 0, args);
        }

        void saveAndCloseWriter(Dictionary<string,string> tagsComm, Discipline disc)
        {
            XStorable stor = xComp as XStorable;

            string fileName = string.Join("_", "РПД", tagsComm["<YEAROFENTRANCE>"],
                tagsComm["<SPECIALIZATION>"].Substring(0, 8), tagsComm["<PROFILEABBR>"].ToLowerInvariant(),
                tagsComm["<FORM>"][0], disc.Code, disc.Abbrevation);

            Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory, "Output"));
            string fullPath = Path.Combine(Environment.CurrentDirectory, "Output", fileName + ".odt");
            fullPath = fullPath.Replace('\\', '/');
            PropertyValue[] args = new PropertyValue[0];

            stor.storeToURL("file:///" + fullPath, args);
            (xComp as XCloseable).close(true);
        }

        void replaceCompetentions(Discipline disc)
        {
            XTextContent compsContext = (XTextContent)(xComp as XBookmarksSupplier)
                .getBookmarks()
                .getByName("ВсеКомпетенции")
                .Value;
            XTextRange compsRange = compsContext.getAnchor();

            StringBuilder comps = new StringBuilder();
            foreach (var kv in disc.Competentions)
            {
                comps.Append($"{kv.Key} – {kv.Value}\n");
            }

            compsRange.setString(comps.ToString());
        }

        void formatCompTable(Discipline disc)
        {
            XTextTable table = (XTextTable)(xComp as XTextTablesSupplier)
                .getTextTables()
                .getByName("Таблица9").Value;

            table.getRows().insertByIndex(2, disc.Competentions.Count);

            int i = 3;
            foreach (var kv in disc.Competentions)
            {
                (table.getCellByName($"A{i}") as unoidl.com.sun.star.text.XText)
                    .setString((i - 2).ToString());
                (table.getCellByName($"B{i}") as unoidl.com.sun.star.text.XText)
                    .setString(kv.Key);
                (table.getCellByName($"C{i}") as unoidl.com.sun.star.text.XText)
                    .setString(kv.Value);

                i++;
            }
        }

        void pasteInCellWorkInfo(int row, int column, int sem, XTextTable tb, WorkInfo wi)
        {
            int hours;
            string col = ((char)(column + 64)).ToString();
            if (wi == null || (hours = wi.HoursOnSemester(sem)) == 0)
            {
                (tb.getCellByName($"{col}{row}") as unoidl.com.sun.star.text.XText)
                    .setString("-");
            }
            else
            {
                (tb.getCellByName($"{col}{row}") as unoidl.com.sun.star.text.XText)
                    .setString(hours.ToString());
            }
        }

        void formatTrudTable(Discipline disc)
        {
            XTextTable table = (XTextTable)(xComp as XTextTablesSupplier)
                .getTextTables()
                .getByName("Таблица10").Value;

            int semCount = disc.Semester.Semesters.Count();
            for (int i = 3; i <= table.getRows().getCount(); i++)
            {
                short newCells = (short)(semCount - 1);
                if (newCells > 0)
                {
                    XTextTableCursor curs = table.createCursorByCellName($"C{i}");
                    curs.splitRange(newCells, false);
                }
            }

            // Колонки с семестрами
            for (int j = 3; j <= disc.Semester.Semesters.Count() + 2; j++)
            {
                int currSem = disc.Semester.Semesters[j - 3];
                string col = ((char)(j + 64)).ToString();
                (table.getCellByName($"{col}{3}") as unoidl.com.sun.star.text.XText)
                    .setString(currSem.ToString());

                // Лекции
                pasteInCellWorkInfo(5, j, currSem, table, disc.Lectures);

                // Лабы
                pasteInCellWorkInfo(6, j, currSem, table, disc.Laboratory);
                pasteInCellWorkInfo(7, j, currSem, table, disc.Laboratory);

                // Практики
                pasteInCellWorkInfo(8, j, currSem, table, disc.Practice);
                pasteInCellWorkInfo(9, j, currSem, table, disc.Practice);

                // Сам. Работы
                pasteInCellWorkInfo(10, j, currSem, table, disc.Independent);
            }
        }

        void formatDiscThemes(Discipline disc)
        {
            XTextTable table = (XTextTable)(xComp as XTextTablesSupplier)
                .getTextTables()
                .getByName("Таблица12").Value;

            for (int i = 2; i <= table.getRows().getCount(); i++)
            {
                short newCells = (short)(disc.Competentions.Count - 1);
                if (newCells > 0)
                {
                    XTextTableCursor curs = table.createCursorByCellName($"B{i}");
                    curs.splitRange(newCells, false);
                }
            }

            int columnCount = 2;
            foreach (var comp in disc.Competentions.Keys)
            {
                string col = ((char)(columnCount + 64)).ToString();
                (table.getCellByName($"{col}{2}") as unoidl.com.sun.star.text.XText)
                    .setString(comp);

                columnCount++;
            }
        }

        public void GenerateDocs(DocAttributes attrs, string pathToTemplate)
        {
            FileInfo templ = new FileInfo(pathToTemplate);
            Dictionary<string, string> tagsCommon = new Dictionary<string, string>() {
                        { "<FACULTY>",          attrs.Faculty                   },
                        { "<DEPARTMENT>",       attrs.Departament               },
                        { "<SPECIALIZATION>",   attrs.Specialization            },
                        { "<PROFILE>",          attrs.Profile                   },
                        { "<PROFILEABBR>",      attrs.ProfileAbbrevation        },
                        { "<EDUCATIONLEVEL>",   attrs.EducationLevel            },
                        { "<FORM>",             attrs.EducationType             },
                        { "<YEAROF>",           attrs.YearOfEntrance.ToString() },
                        { "<GRADUATIONLEVEL>",  attrs.GraduationLevel           },
                        { "<YEAROFENTRANCE>",   attrs.YearOfEntrance.ToString() },
            };

            int count = attrs.Disciplines.Count();
            for (int i = 0; i < count; i++)
            {
                loadWriter(pathToTemplate);

                int totalh = getTotalDisc(attrs.Disciplines[i]);
                string totalle = attrs.Disciplines[i].Lectures == null ? "-" : attrs.Disciplines[i].Lectures.Total.ToString();
                string totalpr = attrs.Disciplines[i].Practice == null ? "-" : attrs.Disciplines[i].Practice.Total.ToString();
                string totalla = attrs.Disciplines[i].Laboratory == null ? "-" : attrs.Disciplines[i].Laboratory.Total.ToString();
                string totalin = attrs.Disciplines[i].Independent == null ? "-" : attrs.Disciplines[i].Independent.Total.ToString();

                Dictionary<string, string> tagsDiscipline = new Dictionary<string, string>() {
                    { "<DISCIPLINE>", attrs.Disciplines[i].Name },
                    { "<TOTALH>",  totalh.ToString() },
                    { "<LECTURESH>", totalle },
                    { "<PRACTICEH>", totalpr },
                    { "<LABORATORYH>", totalla },
                    { "<INDEPENDENTH>", totalin },
                    { "<COURSES>", formatSemArray(attrs.Disciplines[i].Semester.Courses) },
                    { "<SEMESTERS>", formatSemArray(attrs.Disciplines[i].Semester.Semesters) },
                    { "<TOTALCU>", (totalh / 36).ToString() },
                    { "<ACCREDITATION>", formatAttestation(attrs.Disciplines[i].Exam, attrs.Disciplines[i].Credits,
                        attrs.Disciplines[i].RatedCredits) },
                };

                foreach (var tag in tagsCommon)
                    findReplace(tag.Key, tag.Value);

                foreach (var tag in tagsDiscipline)
                    findReplace(tag.Key, tag.Value);

                replaceCompetentions(attrs.Disciplines[i]);
                formatCompTable(attrs.Disciplines[i]);
                formatTrudTable(attrs.Disciplines[i]);
                formatDiscThemes(attrs.Disciplines[i]);

                saveAndCloseWriter(tagsCommon, attrs.Disciplines[i]);
            };
        }

        public void Dispose()
        {
            (xContext as XDesktop)?.terminate();
        }
    }
}
