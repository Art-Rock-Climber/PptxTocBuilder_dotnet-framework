using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace TocBuilder_dotnet_framework.Models
{
    public class SlideItem : INotifyPropertyChanged 
    {
        private bool _isSelected = true; 
        private int _number; 
        private BitmapImage _thumbnail; 

        public int Number { get => _number; set { _number = value; OnPropertyChanged(); } }

        public BitmapImage Thumbnail 
        { 
            get => _thumbnail; 
            set { _thumbnail = value; OnPropertyChanged(); } 
        }

        public bool IsSelected 
        { 
            get => _isSelected; 
            set { _isSelected = value; OnPropertyChanged(); } 
        }

        public event PropertyChangedEventHandler PropertyChanged; 

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null) 
        { 
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName)); 
        }
    }

}
