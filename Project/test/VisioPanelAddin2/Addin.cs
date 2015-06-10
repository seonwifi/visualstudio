using System.Drawing;
using System.Windows.Forms;
using VisioPanelAddin2.Properties;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioPanelAddin2
{
    public partial class Addin
    {
        public Visio.Application Application { get; set; }

        /// <summary>
        /// Callback called by the UI manager when user clicks a button
        /// Should do something meaninful wehn corresponding action is called.
        /// </summary>
        public void OnCommand(string commandId)
        {
            switch (commandId)
            {
                case "Command1":
                    MessageBox.Show(commandId);
                    return;

                case "Command2":
                    MessageBox.Show(commandId);
                    return;

                case "TogglePanel":
                    TogglePanel();
                    return;
            }
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command shoudl be enabled in the user interface.
        /// By default, all commands are enabled.
        /// </summary>
        public bool IsCommandEnabled(string commandId)
        {
            switch (commandId)
            {
                case "Command1":    // make command1 always enabled
                    return true;

                case "Command2":    // make command2 enabled only if a drawing is opened
                    return Application != null && Application.ActiveWindow != null;

                case "TogglePanel": // make panel enabled only if we have an open drawing.
                    return IsPanelEnabled();

                default:
                    return true;
            }
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
        /// </summary>
        public bool IsCommandChecked(string command)
        {

            if (command == "TogglePanel")
                return IsPanelVisible();

            return false;
        }
        /// <summary>
        /// Callback called by UI manager.
        /// Returns a label associated with given command.
        /// We assume for simplicity taht command labels are named simply named as [commandId]_Label (see resources)
        /// </summary>
        public string GetCommandLabel(string command)
        {
            return Resources.ResourceManager.GetString(command + "_Label");
        }

        /// <summary>
        /// Returns a icon associated with given command.
        /// We assume for simplicity that icon ids are named after command commandId.
        /// </summary>
        public Icon GetCommandIcon(string command)
        {
            return (Icon)Resources.ResourceManager.GetObject(command);
        }

        #region Panel
        private void TogglePanel()
        {
            _panelManager.TogglePanel(Application.ActiveWindow);
        }

        private bool IsPanelEnabled()
        {
            return Application != null && Application.ActiveWindow != null;
        }

        private bool IsPanelVisible()
        {
            return Application != null && _panelManager.IsPanelOpened(Application.ActiveWindow);
        }

        private PanelManager _panelManager;
        #endregion

        internal void Startup(object application)
        {
            Application = (Visio.Application)application;

            _panelManager = new PanelManager(this);
        }

        internal void Shutdown()
        {
            _panelManager.Dispose();
        }

        internal void UpdateUI()
        {
            UpdateRibbon();
        }
    }
}
