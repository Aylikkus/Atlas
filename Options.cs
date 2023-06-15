using Atlas.Data;
using Atlas.Interops;
using Atlas.Interops.LibreOffice;
using Atlas.Interops.Office;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Xml;
using System.Xml.Serialization;

namespace Atlas
{
    [DataContract]
    class Options : INotifyPropertyChanged
    {
        [DataMember]
        private Person person;

        private IDocGenerator docGenerator;
        private IDocReader docReader;

        private Options() 
        { 
            docGenerator = new WriterGenerator();
            docReader = new CalcReader();
            person = new Person();
        }
        private static Options instance;

        private static DataContractSerializer serializer = new DataContractSerializer(
                typeof(Options), new List<Type> {
                typeof(CalcReader),
                typeof(WriterGenerator),
                typeof(ExcelReader),
                typeof(WordGenerator),
                typeof(Person),
            });

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        [DataMember]
        public IDocGenerator DocGenerator 
        {
            get { return docGenerator; }
            set
            {
                docGenerator = value;
                OnPropertyChanged();
            }
        }

        [DataMember]
        public IDocReader DocReader
        {
            get { return docReader; }
            set
            {
                docReader = value;
                OnPropertyChanged();
            }
        }

        
        public string Name
        {
            get { return person.Name; }
            set 
            { 
                person.Name = value; 
                OnPropertyChanged();
            }
        }

        public string Surname
        {
            get { return person.Surname; }
            set
            {
                person.Surname = value;
                OnPropertyChanged();
            }
        }

        public string Patronymic
        {
            get { return person.Patronymic; }
            set
            {
                person.Patronymic = value;
                OnPropertyChanged();
            }
        }

        public string Degree
        {
            get { return person.AcademicDegree; }
            set
            {
                person.AcademicDegree = value;
                OnPropertyChanged();
            }
        }

        public string Title
        {
            get { return person.AcademicTitle; }
            set
            {
                person.AcademicTitle = value;
                OnPropertyChanged();
            }
        }

        public string Job
        {
            get { return person.JobTitle; }
            set
            {
                person.JobTitle = value;
                OnPropertyChanged();
            }
        }

        [XmlIgnore]
        public static Options Instance
        {
            get
            {
                if (instance != null)
                {
                    return instance;
                }

                try
                {
                    instance = Load();
                }
                catch
                {
                    instance = new Options();
                    Save();
                }

                return instance;
            }
        }

        public static void Save()
        {
            FileStream fs = File.Open(Path.Combine(Directory.GetCurrentDirectory(), "options.xml"),
                FileMode.Create, FileAccess.Write);
            if (instance == null) { return; }

            serializer.WriteObject(fs, instance);
            fs.Close();
        }

        public static Options Load()
        {
            FileStream reader = new FileStream(
                Path.Combine(Directory.GetCurrentDirectory(), "options.xml"),
                FileMode.Open, FileAccess.Read);
            XmlDictionaryReader xdr = XmlDictionaryReader.CreateTextReader(reader, new XmlDictionaryReaderQuotas());
            Options options = (Options)serializer.ReadObject(xdr, true);
            reader.Close();
            xdr.Close();
            return options;
        }
    }
}
