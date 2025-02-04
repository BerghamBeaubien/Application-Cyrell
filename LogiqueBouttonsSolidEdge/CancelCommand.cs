using System.Windows.Forms;

namespace Application_Cyrell.LogiqueBouttonsSolidEdge
{
    public class CancelCommand : IButtonManager
    {
        private Form _form;

        public CancelCommand(Form form)
        {
            _form = form;
        }

        public void Execute()
        {
            _form.Close();
        }
    }
}