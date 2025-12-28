using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using TocBuilder_dotnet_framework.Models;
using TocBuilder_dotnet_framework.Services;

namespace TocBuilder_dotnet_framework.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private ThumbnailService _thumbnailService;
        private TocGeneratorService _tocService;

        private string _filePath;
        private int _columns = 3;
        private int _margin = 20;
        private string _status;
        private bool _isBusy;
        private bool _useFixedColumns = false;

        private double _previewCanvasWidth = LayoutConstants.DefaultSlideWidth;
        private double _previewCanvasHeight = LayoutConstants.DefaultSlideHeight;

        public double PreviewCanvasWidth
        {
            get => _previewCanvasWidth;
            set { _previewCanvasWidth = value; OnPropertyChanged(); }
        }

        public double PreviewCanvasHeight
        {
            get => _previewCanvasHeight;
            set { _previewCanvasHeight = value; OnPropertyChanged(); }
        }

        private float _actualSlideWidth = LayoutConstants.DefaultSlideWidth;
        private float _actualSlideHeight = LayoutConstants.DefaultSlideHeight;

        public float ActualSlideWidth => _actualSlideWidth;
        public float ActualSlideHeight => _actualSlideHeight;

        public string FilePath
        {
            get => _filePath;
            set
            {
                if (_filePath != value)
                {
                    _filePath = value;
                    OnPropertyChanged();
                    LoadSlides();
                    UpdateCanGenerate();
                }
            }
        }

        public int Columns
        {
            get => _columns;
            set { _columns = value; OnPropertyChanged(); if (UseFixedColumns) UpdatePreview(); }
        }

        public int Margin
        {
            get => _margin;
            set { _margin = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public bool UseFixedColumns
        {
            get => _useFixedColumns;
            set { _useFixedColumns = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        public bool IsBusy
        {
            get => _isBusy;
            set { _isBusy = value; OnPropertyChanged(); UpdateCanGenerate(); }
        }

        public bool CanGenerate => !IsBusy && Slides.Any(s => s.IsSelected);

        public ObservableCollection<SlideItem> Slides { get; } = new ObservableCollection<SlideItem>();
        public ObservableCollection<PreviewItem> PreviewItems { get; } = new ObservableCollection<PreviewItem>();

        public ICommand BrowseCommand { get; }
        public ICommand GenerateCommand { get; }
        public ICommand SelectAllCommand { get; }
        public ICommand DeselectAllCommand { get; }

        public MainViewModel()
        {
            _thumbnailService = new ThumbnailService();
            _tocService = new TocGeneratorService();

            Status = "Выберите презентацию";

            BrowseCommand = new RelayCommand(BrowseFile);
            GenerateCommand = new RelayCommand(GenerateToc, () => CanGenerate);
            SelectAllCommand = new RelayCommand(() => SelectAll(true));
            DeselectAllCommand = new RelayCommand(() => SelectAll(false));

            PropertyChanged += MainViewModel_PropertyChanged;
        }

        private void MainViewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(Margin) || e.PropertyName == nameof(Columns))
            {
                UpdatePreview();
            }
        }

        private void Slide_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SlideItem.IsSelected))
            {
                UpdatePreview();
                UpdateCanGenerate();
            }
        }

        private void BrowseFile()
        {
            var dlg = new OpenFileDialog { Filter = "PowerPoint презентации|*.pptx;*.ppt", Title = "Выберите презентацию" };
            if (dlg.ShowDialog() == true)
            {
                FilePath = dlg.FileName;
            }
        }

        private void LoadSlides()
        {
            Slides.Clear();
            if (string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath)) return;

            IsBusy = true;
            Status = "Загрузка слайдов...";

            try
            {
                (_actualSlideWidth, _actualSlideHeight) = _thumbnailService.GetSlideDimensions(FilePath);
                var slides = _thumbnailService.GetSlides(FilePath);
                foreach (var slide in slides)
                {
                    slide.PropertyChanged += Slide_PropertyChanged;
                    Slides.Add(slide);
                }

                Status = $"Загружено {Slides.Count} слайдов";
                UpdatePreview();
                UpdateCanGenerate();
            }
            catch (Exception ex)
            {
                Status = $"Ошибка: {ex.Message}";
                MessageBox.Show($"Не удалось загрузить презентацию:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void UpdatePreview()
        {
            var selectedSlides = Slides.Where(s => s.IsSelected).ToList();
            if (!selectedSlides.Any())
            {
                PreviewItems.Clear();
                return;
            }

            int columnsForCalc = UseFixedColumns ? Columns : -1;
            var layoutInfo = LayoutCalculatorService.CalculateOptimalLayout(
                    selectedSlides.Count,
                    Margin,
                    columnsForCalc,
                    _actualSlideWidth,
                    _actualSlideHeight);
            var previewItems = LayoutCalculatorService.GeneratePreviewItems(
                    selectedSlides,
                    columnsForCalc,
                    Margin,
                    _actualSlideWidth,
                    _actualSlideHeight);

            var (canvasWidth, canvasHeight) = LayoutCalculatorService.CalculateCanvasSize(previewItems);

            Application.Current.Dispatcher.Invoke(() =>
            {
                PreviewItems.Clear();
                PreviewCanvasWidth = canvasWidth;
                PreviewCanvasHeight = canvasHeight;
                foreach (var item in previewItems)
                    PreviewItems.Add(item);
            });
        }



        private void GenerateToc()
        {
            if (!Slides.Any(s => s.IsSelected)) return;

            IsBusy = true;
            Status = "Создание оглавления...";

            try
            {
                var selectedSlides = Slides.Where(s => s.IsSelected).ToList();
                string outputPath = _tocService.CreateTableOfContents(FilePath, selectedSlides, Columns, Margin);

                Status = $"✅ Готово! Файл сохранён: {Path.GetFileName(outputPath)}";

                if (MessageBox.Show($"Оглавление создано!\n\nФайл: {outputPath}\n\nОткрыть презентацию?", "Готово", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Status = $"❌ Ошибка: {ex.Message}";
                MessageBox.Show($"Не удалось создать оглавление:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                UpdateCanGenerate();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void SelectAll(bool select)
        {
            foreach (var slide in Slides) slide.IsSelected = select;
            UpdateCanGenerate();
        }

        private void UpdateCanGenerate()
        {
            OnPropertyChanged(nameof(CanGenerate));
            if (GenerateCommand is RelayCommand cmd) cmd.RaiseCanExecuteChanged();
        }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        #endregion
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;
        public event EventHandler CanExecuteChanged;

        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) => _canExecute?.Invoke() ?? true;
        public void Execute(object parameter) => _execute();
        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}