using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

namespace Application_Cyrell.LogiqueBouttonsSolidEdge
{
    public interface IButtonManager
    {
        void Execute();
    }

    public abstract class SolidEdgeCommandBase : IButtonManager
    {
        protected TextBox _textBoxFolderPath;
        protected ListBox _listBoxDxfFiles;

        public SolidEdgeCommandBase(TextBox textBoxFolderPath, ListBox listBoxDxfFiles)
        {
            _textBoxFolderPath = textBoxFolderPath;
            _listBoxDxfFiles = listBoxDxfFiles;
        }

        public abstract void Execute();
    }
}
