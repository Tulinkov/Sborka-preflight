using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Corel.Interop.VGCore;
using SborkaPreflight.Model;

namespace SborkaPreflight.ViewModel
{
    class MainUserControlWM
    {
        public CorelDockerModel CorelDocker { get; set; }
        public ICommand BtnOrderList_ClickCommand { get; set; }
        public ICommand BtnToPDF_ClickCommand { get; set; }
        public ICommand BtnToCDR_ClickCommand { get; set; }
        public ICommand BtnToPrepress_ClickCommand { get; set; }
        public ICommand BtnToUser_ClickCommand { get; set; }
        public ICommand BtnGuides_ClickCommand { get; set; }
        public ICommand BtnFitToPage_ClickCommand { get; set; }
        public ICommand BtnNewOrder_ClickCommand { get; set; }
        public ICommand BtnSetting_ClickCommand { get; set; }
        public ICommand Copy { get; set; }
        public ICommand BtnUpdate_ClickCommand { get; set; }
        public ICommand BtnSideOpen_ClickCommand { get; set; }
        public ICommand BtnDrawCircle_ClickCommand { get; set; }
        public ICommand BtnFixCutline_ClickCommand { get; set; }

        public MainUserControlWM(Corel.Interop.VGCore.Application app)
        {
            CorelDocker = new CorelDockerModel(app);
            BtnOrderList_ClickCommand = new Command(arg => CorelDocker.CreateListOfOrder());
            BtnToPDF_ClickCommand = new Command(arg => CorelDocker.PublishToPDF());
            BtnToCDR_ClickCommand = new Command(arg => CorelDocker.SaveToV17());
            BtnToPrepress_ClickCommand = new Command(arg => CorelDocker.SendToPrepress());
            BtnToUser_ClickCommand = new Command(arg => CorelDocker.ReturnToUser());
            BtnGuides_ClickCommand = new Command(arg => CorelDocker.MakeOrDeleteGuides());
            BtnFitToPage_ClickCommand = new Command(arg => CorelDocker.FitToPage());
            BtnNewOrder_ClickCommand = new Command(arg => CorelDocker.ProcessOrder(arg));
            BtnSetting_ClickCommand = new Command(arg => CorelDocker.SettingWindow());
            Copy = new Command(arg => CorelDocker.Copy(arg));
            BtnUpdate_ClickCommand = new Command(arg => CorelDocker.Update());
            BtnSideOpen_ClickCommand = new Command(arg => CorelDocker.SideOpen(arg));
            BtnDrawCircle_ClickCommand = new Command(arg => CorelDocker.DrawCircle());
            BtnFixCutline_ClickCommand = new Command(arg => CorelDocker.FixCutline());
        }
    }
}
