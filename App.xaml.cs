using System;
using Microsoft.UI.Xaml;

namespace App1
{
    public partial class App : Application
    {
        private Window? _window;
            
        public App()
        {
   
            InitializeComponent();
        }

  
        /// <param name="args"
        
        protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            _window = new MainWindow();
            _window.Activate();
        }
    }
}
