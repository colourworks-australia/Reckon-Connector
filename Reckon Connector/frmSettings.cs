using System;
using System.Windows.Forms;
using ReckonDesktop.Repository;
using ReckonDesktop.Model;

namespace Reckon_Connector
{
    public enum FormMode
    {
        Add,
        Edit
    }

    public partial class frmSettings : Form
    {
        private readonly UnitOfWork _unitOfWork;
        private readonly FormMode _formMode;
        private readonly Settings _settings;
        private readonly int _id;

        public frmSettings(FormMode formMode, UnitOfWork uow, int id = 1)
        {
            InitializeComponent();
            _formMode = formMode;
            _unitOfWork = uow;
            _id = id;

            if (id != 0)
            {
                _settings = _unitOfWork.SettingsRepository.GetByID(_id);
            }

            switch (formMode)
            {
                case FormMode.Add:
                    cmdAdd.Text = "Add";
                    break;
                case FormMode.Edit:
                    cmdAdd.Text = "Update";
                    break;
            }

        }

        private bool Validate()
        {
            //if (txtMappingSetName.Text.Trim() != "")
            //{
            //    if (cmbDSN.Text != "")
            //    {
            //        if (txtQuery.Text != "")
            //        {
            //            if (CheckDSNConnection())
            //            {
            //                if (cmbProject.Text != "")
            //                {
            //                    return true;
            //                }
            //                else
            //                {
            //                    MessageBox.Show("Project has not been set.", "Issue", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //                    return false;
            //                }
            //            }
            //            else
            //            {
            //                MessageBox.Show("Connection to DSN failed. Please correct the DSN or Query string.", "Issue", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //                return false;
            //            }
            //        }
            //        else
            //        {
            //            MessageBox.Show("Query string has not been set.", "Issue", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            return false;
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show("DSN has not been set.", "Issue", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        return false;
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Mapping Set Name cannot be blank.", "Issue", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}
            return true;
        }

        private void cmdAdd_Click(object sender, EventArgs e)
        {

            if (Validate())
            {
                switch (_formMode)
                {
                    case FormMode.Add:
                        var newDestination = new Settings();

                        newDestination.UseHosted = checkBox1.Checked;
                        newDestination.ConnectionString = txtConnectionString.Text;
                        newDestination.MyobUser = txtMyobUser.Text;
                        newDestination.MyobPassword = txtMyobPassword.Text;
                        newDestination.AutofileEndpoint = txtAutofileEndpoint.Text;
                        newDestination.AutofileUser = txtAutofileUser.Text;
                        newDestination.AutofilePassword = txtAutofilePassword.Text;
                      //  newDestination.LogFilePath = txtLogFilePath.Text;

                        _unitOfWork.SettingsRepository.Insert(newDestination);
                        _unitOfWork.Save();

                        break;
                    case FormMode.Edit:
                        var existingId = txtId.Text;
                        var intId = UtilityHelper.IntParseToDefaultValue(existingId, 1);
                        var existingDestination = _unitOfWork.SettingsRepository.GetByID(intId);
                        existingDestination.UseHosted = checkBox1.Checked;
                        existingDestination.ConnectionString = txtConnectionString.Text;
                        existingDestination.MyobUser = txtMyobUser.Text;
                        existingDestination.MyobPassword = txtMyobPassword.Text;
                        existingDestination.AutofileEndpoint = txtAutofileEndpoint.Text;
                        existingDestination.AutofileUser = txtAutofileUser.Text;
                        existingDestination.AutofilePassword = txtAutofilePassword.Text;
                      //  existingDestination.LogFilePath = txtLogFilePath.Text;

                        _unitOfWork.SettingsRepository.Detach(existingDestination);
                        _unitOfWork.SettingsRepository.Update(existingDestination);
                        _unitOfWork.Save();
                        MessageBox.Show("Connection updated", "Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        break;
                }
                Close();
            }
        }

        private void frmDestination_Load(object sender, EventArgs e)
        {
            checkBox1.Visible = true;
            txtConnectionString.Visible = true;
            txtMyobUser.Visible = true;
            txtMyobPassword.Visible = true;
            txtAutofileEndpoint.Visible = true;
            txtAutofileUser.Visible = true;
            txtAutofilePassword.Visible = true;
           // txtLogFilePath.Visible = true;
            if (_settings != null)
            {
                if (_settings.UseHosted != null && _settings.UseHosted.Value == true) { checkBox1.Checked = true;}
                txtConnectionString.Text = _settings.ConnectionString;
                txtMyobUser.Text = _settings.MyobUser;
                txtMyobPassword.Text = _settings.MyobPassword;
                txtAutofileEndpoint.Text = _settings.AutofileEndpoint;
                txtAutofileUser.Text = _settings.AutofileUser;
                txtAutofilePassword.Text = _settings.AutofilePassword;
                txtId.Text = _settings.Id.ToString();
                // txtLogFilePath.Text = _settings.LogFilePath;
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void cmdAdd1_Click(object sender, EventArgs e)
        {
            var newDestination = new Settings();

            newDestination.UseHosted = checkBox1.Checked;
            newDestination.ConnectionString = txtConnectionString.Text;
            newDestination.MyobUser = txtMyobUser.Text;
            newDestination.MyobPassword = txtMyobPassword.Text;
            newDestination.AutofileEndpoint = txtAutofileEndpoint.Text;
            newDestination.AutofileUser = txtAutofileUser.Text;
            newDestination.AutofilePassword = txtAutofilePassword.Text;
            //  newDestination.LogFilePath = txtLogFilePath.Text;

            _unitOfWork.SettingsRepository.Insert(newDestination);
            _unitOfWork.Save();
            MessageBox.Show("New connection added", "Added", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void cmdDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure that you want to delete this connection?", "Delete?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var existingId = txtId.Text;
                var intId = UtilityHelper.IntParseToDefaultValue(existingId, 1);
                var existingDestination = _unitOfWork.SettingsRepository.GetByID(intId);

                _unitOfWork.SettingsRepository.Delete(existingDestination);
                _unitOfWork.Save();
                MessageBox.Show("Connection deleted", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
            }
        }
    }
}
