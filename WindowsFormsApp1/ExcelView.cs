using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class ExcelView : Form
    {
        public ExcelView( )
        {
            ce = new CreateExcell();
            InitializeComponent( );
        }

        private void searchFileNazwisko_Click( object sender, EventArgs e )
        {
            searchFileNaz.ShowDialog( );
        }

        private void searchFileStanowisko_Click( object sender, EventArgs e )
        {
            searchFileSta.ShowDialog( );
        }

        private void searchFileNaz_FileOk( object sender, CancelEventArgs e )
        {
            pathNazwisko.Text = searchFileNaz.FileName;
        }

        private void searchFileSta_FileOk( object sender, CancelEventArgs e )
        {
            pathStanowisko.Text = searchFileSta.FileName;
        }

        private void buttonSearchToFile_Click( object sender, EventArgs e )
        {
            searchToFile.ShowDialog( );
        }

        private void searchToFile_FileOk( object sender, CancelEventArgs e )
        {
            pathToFile.Text = searchToFile.FileName;
        }

        private void butCreateSheet_Click( object sender, EventArgs e )
        {
            //ce = new CreateExcell( );
            ce.OpenFiles( pathNazwisko.Text, pathStanowisko.Text, pathToFile.Text );
            if ( ce.IsOpenNaz )
            {
                isOpenNazwisko.ForeColor = Color.Green;
                isOpenNazwisko.Text = "Successful";
                isOpenNazwisko.Visible = true;
            }
            else
            {
                isOpenNazwisko.ForeColor = Color.Red;
                isOpenNazwisko.Text = "Error operation";
                isOpenNazwisko.Visible = true;
            }

            if ( ce.IsOpenStan )
            {
                isOpenStan.ForeColor = Color.Green;
                isOpenStan.Text = "Successful";
                isOpenStan.Visible = true;
            }
            else
            {
                isOpenStan.ForeColor = Color.Red;
                isOpenStan.Text = "Error operation";
                isOpenStan.Visible = true;
            }
            if ( ce.IsOpenCreateFile )
            {
                isOpenToFile.ForeColor = Color.Green;
                isOpenToFile.Text = "Successful";
                isOpenToFile.Visible = true;
            }
            else
            {
                isOpenToFile.ForeColor = Color.Red;
                isOpenToFile.Text = "Error operation";
                isOpenToFile.Visible = true;

            }
        }

        private void passNazTurn_CheckedChanged( object sender, EventArgs e )
        {
            passNazwisko.Visible = passNazTurn.Checked;
        }

        private void passStanTurn_CheckedChanged( object sender, EventArgs e )
        {
            passStanowisco.Visible = passStanTurn.Checked;
        }

        private void passToFileTurn_CheckedChanged( object sender, EventArgs e )
        {
            PassToFile.Visible = passToFileTurn.Checked;
        }

        private void passNazwisko_TextChanged( object sender, EventArgs e )
        {
            ce.passwordNazwisko = passNazwisko.Text;
        }

        private void passStanowisco_TextChanged( object sender, EventArgs e )
        {
            ce.passwordStanowisko = passNazwisko.Text;
        }

        private void PassToFile_TextChanged( object sender, EventArgs e )
        {
            ce.passwordCreateFile = PassToFile.Text;
        }

        private CreateExcell ce;
    }
}
