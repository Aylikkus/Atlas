using Atlas.Data;
using Atlas.Interops.LibreOffice;
using Atlas.Interops.Office;
using Microsoft.Office.Core;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Shapes;

namespace Atlas
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        DocAttributes attributes = new DocAttributes();
        CancellationTokenSource tokenSource = new CancellationTokenSource();
        Options options = Options.Instance;

        public MainWindow()
        {
            InitializeComponent();
            if (options.Name == null || options.Surname == null)
            {
                LoginWindow loginWindow = new LoginWindow();
                loginWindow.Show();
                Hide();
                loginWindow.Closed += (s, e) =>
                {
                    Show();
                };
            }

            personBlock.Text = string.Join(" ", options.Surname, options.Name, options.Patronymic);
            if (options.DocReader is CalcReader)
            {
                readerBox.SelectedIndex = 0;
            }
            else if (options.DocReader is ExcelReader)
            {
                readerBox.SelectedIndex = 1;
            }

            if (options.DocGenerator is WriterGenerator)
            {
                generatorBox.SelectedIndex = 0;
            }
            else if (options.DocGenerator is WordGenerator)
            {
                generatorBox.SelectedIndex = 1;
            }
        }

        private void saveBtn_Click(object sender, RoutedEventArgs e)
        {
            Options.Save();
        }

        private void personBlock_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            LoginWindow loginWindow = new LoginWindow();
            loginWindow.Show();
            Hide();
            loginWindow.Closed += (s, ea) =>
            {
                personBlock.Text = string.Join(" ", options.Surname, options.Name, options.Patronymic);
                Show();
            };
        }

        private void readerBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            switch (readerBox.SelectedIndex) 
            {
                case 0:
                    options.DocReader = new CalcReader();
                    break;
                case 1:
                    options.DocReader = new ExcelReader();
                    break;
            }
        }

        private void generatorBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            switch (generatorBox.SelectedIndex)
            {
                case 0:
                    options.DocGenerator = new WriterGenerator();
                    break;
                case 1:
                    options.DocGenerator = new WordGenerator();
                    break;
            }
        }

        private void readBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == false) return;
            Task task = new Task(() =>
            {
                try
                {
                    attributes = options.DocReader.PullAttributes(dialog.FileName);
                    StringBuilder statsBuilder = new StringBuilder();
                    statsBuilder.AppendLine("Специализация: " + attributes.Specialization);
                    statsBuilder.AppendLine("Направленность: " + attributes.Profile);
                    statsBuilder.AppendLine("Аббревиатура: " + attributes.ProfileAbbrevation);
                    statsBuilder.AppendLine("Кафедра: " + attributes.Departament);
                    statsBuilder.AppendLine("Факультет: " + attributes.Faculty);
                    statsBuilder.AppendLine("Количество дисциплин: " + attributes.Disciplines.Count);
                    statsBuilder.AppendLine("Уровень образования: " + attributes.EducationLevel);
                    statsBuilder.AppendLine(attributes.GraduationLevel);
                    statsBuilder.AppendLine("Год набора: " + attributes.YearOfEntrance);
                    statsBuilder.Append("Компетенции:");
                    foreach (var kv in attributes.Competentions)
                    {
                        statsBuilder.Append(" " + kv.Key);
                    }

                    statsBlock.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        statsBlock.Text = statsBuilder.ToString();
                        statsBlock.Visibility = Visibility.Visible;
                    }));
                    generateBtn.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        generateBtn.IsEnabled = true;
                    }));
                }
                catch
                {
                    MessageBox.Show("Проблема при чтении учебного плана", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    statsBlock.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        statsBlock.Text = "Чтение завершилось ошибкой";
                        statsBlock.Visibility = Visibility.Visible;
                    }));
                    return;
                }
            });
            task.Start();
        }

        private void generateBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == false) return;
            stopBtn.IsEnabled = true;
            stopBtn.Visibility = Visibility.Visible;
            tokenSource.Token.Register(() =>
            {
                throw new OperationCanceledException();
            });
            Task task = new Task(() =>
            {
                options.DocGenerator.DisciplineFinished += (i) =>
                {
                    progressBlock.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        progressBlock.Text = "Прогресс: " + i + "/" + attributes.Disciplines.Count;
                    }));
                };
                try
                {
                    Thread thread = new Thread(() =>
                    {
                        options.DocGenerator.GenerateDocs(attributes, dialog.FileName);
                    });
                    thread.Start();
                    while (true)
                    {
                        Thread.Sleep(1000);
                        if (thread.ThreadState == ThreadState.Stopped)
                        {
                            break;
                        }
                        if (tokenSource.IsCancellationRequested)
                        {
                            tokenSource = new CancellationTokenSource();
                            thread.Abort();
                            break;
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Проблема при загрузке шаблона", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                stopBtn.Dispatcher.BeginInvoke(new Action(() =>
                {
                    stopBtn.IsEnabled = false;
                    stopBtn.Visibility = Visibility.Collapsed;
                }));
            }, tokenSource.Token);
            progressBlock.Visibility = Visibility.Visible;
            task.Start();
        }

        private void stopBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                tokenSource.Cancel();
            }
            catch (AggregateException)
            {
                MessageBox.Show("Генерация РПД успешно отменена", "Отмена", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            stopBtn.IsEnabled = false;
            stopBtn.Visibility = Visibility.Collapsed;
        }
    }
}
