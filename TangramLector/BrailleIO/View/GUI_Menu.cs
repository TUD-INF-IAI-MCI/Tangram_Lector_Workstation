
namespace tud.mci.tangram.TangramLector
{
    public class GUI_Menu
    {
        public const int NONE_MENU = 0;
        public const int SAVE_MENU = 1;  // SAVE_MENU
        private int activeMenu;        

        public void resetGuiMenu()
        {
            
            activeMenu = GUI_Menu.NONE_MENU;
        }

        public int ActiveMenu
        {
            get { return activeMenu; }
        }


    }
}
