using System;
using System.Linq;
using System.Windows.Forms;
using ReckonDesktop.Model;
using ReckonDesktop.Repository;

namespace Reckon_Connector
{
    public partial class frmProtect : Form
    {
        private readonly UnitOfWork _unitOfWork;
        public frmProtect(UnitOfWork uow)
        {
            InitializeComponent();
            _unitOfWork = uow;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void cmdAdd_Click(object sender, EventArgs e)
        {
            if (txtPw1.Text == txtPw2.Text)
            {
                var existingDestination = _unitOfWork.SecurityRepository.Get(x => x.Login == "system").FirstOrDefault();
                if (existingDestination != null)
                {
                    existingDestination.Password = txtPw1.Text;
                    existingDestination.Protected = chkEnable.Checked;
                    _unitOfWork.SecurityRepository.Detach(existingDestination);
                    _unitOfWork.SecurityRepository.Update(existingDestination);
                }
                else
                {
                    var newPassword = new Security
                    {
                        Login = "system",
                        Password = txtPw1.Text,
                        Protected = chkEnable.Checked
                    };
                    _unitOfWork.SecurityRepository.Insert(newPassword);
                }
                _unitOfWork.Save();
                Close();
            }
            else
            {
                MessageBox.Show("Passwords do not match.", "Issue", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtPw1.Focus();
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmProtect_Load(object sender, EventArgs e)
        {
            var existingDestination = _unitOfWork.SecurityRepository.Get(x => x.Login == "system").FirstOrDefault();
            if (existingDestination != null)
            {
                txtPw1.Text = existingDestination.Password;
                txtPw2.Text = existingDestination.Password;
                chkEnable.Checked = existingDestination.Protected;
            }
        }
    }
}
