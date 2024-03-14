
namespace WindowsFormsApp1
{
    partial class ExcelView
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose( bool disposing )
        {
            if ( disposing && ( components != null ) )
            {
                components.Dispose( );
            }
            base.Dispose( disposing );
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent( )
        {
            this.butSearchFileNazwisko = new System.Windows.Forms.Button();
            this.pathNazwisko = new System.Windows.Forms.Label();
            this.fromName = new System.Windows.Forms.Label();
            this.Stanowisko = new System.Windows.Forms.Label();
            this.butSearchFileStanowisko = new System.Windows.Forms.Button();
            this.pathStanowisko = new System.Windows.Forms.Label();
            this.searchFileNaz = new System.Windows.Forms.OpenFileDialog();
            this.searchFileSta = new System.Windows.Forms.OpenFileDialog();
            this.createToFile = new System.Windows.Forms.Label();
            this.pathToFile = new System.Windows.Forms.Label();
            this.buttonSearchToFile = new System.Windows.Forms.Button();
            this.searchToFile = new System.Windows.Forms.OpenFileDialog();
            this.butCreateSheet = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.isOpenNazwisko = new System.Windows.Forms.Label();
            this.isOpenToFile = new System.Windows.Forms.Label();
            this.isOpenStan = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.passNazTurn = new System.Windows.Forms.CheckBox();
            this.passNazwisko = new System.Windows.Forms.TextBox();
            this.passStanTurn = new System.Windows.Forms.CheckBox();
            this.passToFileTurn = new System.Windows.Forms.CheckBox();
            this.passStanowisco = new System.Windows.Forms.TextBox();
            this.PassToFile = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // butSearchFileNazwisko
            // 
            this.butSearchFileNazwisko.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.butSearchFileNazwisko.Location = new System.Drawing.Point(12, 57);
            this.butSearchFileNazwisko.Name = "butSearchFileNazwisko";
            this.butSearchFileNazwisko.Size = new System.Drawing.Size(616, 39);
            this.butSearchFileNazwisko.TabIndex = 0;
            this.butSearchFileNazwisko.Text = "Search";
            this.butSearchFileNazwisko.UseVisualStyleBackColor = true;
            this.butSearchFileNazwisko.Click += new System.EventHandler(this.searchFileNazwisko_Click);
            // 
            // pathNazwisko
            // 
            this.pathNazwisko.AutoSize = true;
            this.pathNazwisko.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.pathNazwisko.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.pathNazwisko.Location = new System.Drawing.Point(161, 9);
            this.pathNazwisko.Name = "pathNazwisko";
            this.pathNazwisko.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.pathNazwisko.Size = new System.Drawing.Size(43, 20);
            this.pathNazwisko.TabIndex = 1;
            this.pathNazwisko.Text = "Path";
            // 
            // fromName
            // 
            this.fromName.AutoSize = true;
            this.fromName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.fromName.Location = new System.Drawing.Point(12, 9);
            this.fromName.Name = "fromName";
            this.fromName.Size = new System.Drawing.Size(135, 20);
            this.fromName.TabIndex = 2;
            this.fromName.Text = "From Nazwisko: ";
            // 
            // Stanowisko
            // 
            this.Stanowisko.AutoSize = true;
            this.Stanowisko.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Stanowisko.Location = new System.Drawing.Point(12, 118);
            this.Stanowisko.Name = "Stanowisko";
            this.Stanowisko.Size = new System.Drawing.Size(143, 20);
            this.Stanowisko.TabIndex = 3;
            this.Stanowisko.Text = "From Stanowisko ";
            // 
            // butSearchFileStanowisko
            // 
            this.butSearchFileStanowisko.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.butSearchFileStanowisko.Location = new System.Drawing.Point(12, 153);
            this.butSearchFileStanowisko.Name = "butSearchFileStanowisko";
            this.butSearchFileStanowisko.Size = new System.Drawing.Size(616, 39);
            this.butSearchFileStanowisko.TabIndex = 4;
            this.butSearchFileStanowisko.Text = "Search";
            this.butSearchFileStanowisko.UseVisualStyleBackColor = true;
            this.butSearchFileStanowisko.Click += new System.EventHandler(this.searchFileStanowisko_Click);
            // 
            // pathStanowisko
            // 
            this.pathStanowisko.AutoSize = true;
            this.pathStanowisko.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.pathStanowisko.Location = new System.Drawing.Point(161, 118);
            this.pathStanowisko.Name = "pathStanowisko";
            this.pathStanowisko.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.pathStanowisko.Size = new System.Drawing.Size(43, 20);
            this.pathStanowisko.TabIndex = 5;
            this.pathStanowisko.Text = "Path";
            // 
            // searchFileNaz
            // 
            this.searchFileNaz.FileName = "searxhFileNazwisko";
            this.searchFileNaz.FileOk += new System.ComponentModel.CancelEventHandler(this.searchFileNaz_FileOk);
            // 
            // searchFileSta
            // 
            this.searchFileSta.FileName = "searchFileStanowisko";
            this.searchFileSta.FileOk += new System.ComponentModel.CancelEventHandler(this.searchFileSta_FileOk);
            // 
            // createToFile
            // 
            this.createToFile.AutoSize = true;
            this.createToFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.createToFile.Location = new System.Drawing.Point(12, 213);
            this.createToFile.Name = "createToFile";
            this.createToFile.Size = new System.Drawing.Size(55, 20);
            this.createToFile.TabIndex = 6;
            this.createToFile.Text = "To file";
            // 
            // pathToFile
            // 
            this.pathToFile.AutoSize = true;
            this.pathToFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.pathToFile.Location = new System.Drawing.Point(87, 213);
            this.pathToFile.Name = "pathToFile";
            this.pathToFile.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.pathToFile.Size = new System.Drawing.Size(43, 20);
            this.pathToFile.TabIndex = 7;
            this.pathToFile.Text = "Path";
            // 
            // buttonSearchToFile
            // 
            this.buttonSearchToFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonSearchToFile.Location = new System.Drawing.Point(16, 241);
            this.buttonSearchToFile.Name = "buttonSearchToFile";
            this.buttonSearchToFile.Size = new System.Drawing.Size(616, 39);
            this.buttonSearchToFile.TabIndex = 8;
            this.buttonSearchToFile.Text = "Search";
            this.buttonSearchToFile.UseVisualStyleBackColor = true;
            this.buttonSearchToFile.Click += new System.EventHandler(this.buttonSearchToFile_Click);
            // 
            // searchToFile
            // 
            this.searchToFile.FileName = "searchToExcFile";
            this.searchToFile.FileOk += new System.ComponentModel.CancelEventHandler(this.searchToFile_FileOk);
            // 
            // butCreateSheet
            // 
            this.butCreateSheet.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.butCreateSheet.Location = new System.Drawing.Point(16, 319);
            this.butCreateSheet.Name = "butCreateSheet";
            this.butCreateSheet.Size = new System.Drawing.Size(616, 39);
            this.butCreateSheet.TabIndex = 9;
            this.butCreateSheet.Text = "Create sheet";
            this.butCreateSheet.UseVisualStyleBackColor = true;
            this.butCreateSheet.Click += new System.EventHandler(this.butCreateSheet_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(-92, 13);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(98, 21);
            this.checkBox1.TabIndex = 10;
            this.checkBox1.Text = "checkBox1";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // isOpenNazwisko
            // 
            this.isOpenNazwisko.AutoSize = true;
            this.isOpenNazwisko.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.isOpenNazwisko.Location = new System.Drawing.Point(647, 68);
            this.isOpenNazwisko.Name = "isOpenNazwisko";
            this.isOpenNazwisko.Size = new System.Drawing.Size(121, 20);
            this.isOpenNazwisko.TabIndex = 11;
            this.isOpenNazwisko.Text = "Error operation";
            this.isOpenNazwisko.Visible = false;
            // 
            // isOpenToFile
            // 
            this.isOpenToFile.AutoSize = true;
            this.isOpenToFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.isOpenToFile.Location = new System.Drawing.Point(647, 252);
            this.isOpenToFile.Name = "isOpenToFile";
            this.isOpenToFile.Size = new System.Drawing.Size(121, 20);
            this.isOpenToFile.TabIndex = 12;
            this.isOpenToFile.Text = "Error operation";
            this.isOpenToFile.Visible = false;
            // 
            // isOpenStan
            // 
            this.isOpenStan.AutoSize = true;
            this.isOpenStan.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.isOpenStan.Location = new System.Drawing.Point(647, 164);
            this.isOpenStan.Name = "isOpenStan";
            this.isOpenStan.Size = new System.Drawing.Size(121, 20);
            this.isOpenStan.TabIndex = 13;
            this.isOpenStan.Text = "Error operation";
            this.isOpenStan.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(818, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 20);
            this.label1.TabIndex = 14;
            this.label1.Text = "Password";
            // 
            // passNazTurn
            // 
            this.passNazTurn.AutoSize = true;
            this.passNazTurn.Location = new System.Drawing.Point(791, 66);
            this.passNazTurn.Name = "passNazTurn";
            this.passNazTurn.Size = new System.Drawing.Size(18, 17);
            this.passNazTurn.TabIndex = 15;
            this.passNazTurn.UseVisualStyleBackColor = true;
            this.passNazTurn.CheckedChanged += new System.EventHandler(this.passNazTurn_CheckedChanged);
            // 
            // passNazwisko
            // 
            this.passNazwisko.Location = new System.Drawing.Point(822, 65);
            this.passNazwisko.Name = "passNazwisko";
            this.passNazwisko.PasswordChar = '*';
            this.passNazwisko.Size = new System.Drawing.Size(246, 22);
            this.passNazwisko.TabIndex = 16;
            this.passNazwisko.Visible = false;
            this.passNazwisko.TextChanged += new System.EventHandler(this.passNazwisko_TextChanged);
            // 
            // passStanTurn
            // 
            this.passStanTurn.AutoSize = true;
            this.passStanTurn.Location = new System.Drawing.Point(791, 168);
            this.passStanTurn.Name = "passStanTurn";
            this.passStanTurn.Size = new System.Drawing.Size(18, 17);
            this.passStanTurn.TabIndex = 17;
            this.passStanTurn.UseVisualStyleBackColor = true;
            this.passStanTurn.CheckedChanged += new System.EventHandler(this.passStanTurn_CheckedChanged);
            // 
            // passToFileTurn
            // 
            this.passToFileTurn.AutoSize = true;
            this.passToFileTurn.Location = new System.Drawing.Point(790, 256);
            this.passToFileTurn.Name = "passToFileTurn";
            this.passToFileTurn.Size = new System.Drawing.Size(18, 17);
            this.passToFileTurn.TabIndex = 18;
            this.passToFileTurn.UseVisualStyleBackColor = true;
            this.passToFileTurn.CheckedChanged += new System.EventHandler(this.passToFileTurn_CheckedChanged);
            // 
            // passStanowisco
            // 
            this.passStanowisco.Location = new System.Drawing.Point(822, 164);
            this.passStanowisco.Name = "passStanowisco";
            this.passStanowisco.PasswordChar = '*';
            this.passStanowisco.Size = new System.Drawing.Size(246, 22);
            this.passStanowisco.TabIndex = 19;
            this.passStanowisco.Visible = false;
            this.passStanowisco.TextChanged += new System.EventHandler(this.passStanowisco_TextChanged);
            // 
            // PassToFile
            // 
            this.PassToFile.Location = new System.Drawing.Point(822, 255);
            this.PassToFile.Name = "PassToFile";
            this.PassToFile.PasswordChar = '*';
            this.PassToFile.Size = new System.Drawing.Size(246, 22);
            this.PassToFile.TabIndex = 20;
            this.PassToFile.Visible = false;
            this.PassToFile.TextChanged += new System.EventHandler(this.PassToFile_TextChanged);
            // 
            // ExcelView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1185, 493);
            this.Controls.Add(this.PassToFile);
            this.Controls.Add(this.passStanowisco);
            this.Controls.Add(this.passToFileTurn);
            this.Controls.Add(this.passStanTurn);
            this.Controls.Add(this.passNazwisko);
            this.Controls.Add(this.passNazTurn);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.isOpenStan);
            this.Controls.Add(this.isOpenToFile);
            this.Controls.Add(this.isOpenNazwisko);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.butCreateSheet);
            this.Controls.Add(this.buttonSearchToFile);
            this.Controls.Add(this.pathToFile);
            this.Controls.Add(this.createToFile);
            this.Controls.Add(this.pathStanowisko);
            this.Controls.Add(this.butSearchFileStanowisko);
            this.Controls.Add(this.Stanowisko);
            this.Controls.Add(this.fromName);
            this.Controls.Add(this.pathNazwisko);
            this.Controls.Add(this.butSearchFileNazwisko);
            this.Name = "ExcelView";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button butSearchFileNazwisko;
        private System.Windows.Forms.Label pathNazwisko;
        private System.Windows.Forms.Label fromName;
        private System.Windows.Forms.Label Stanowisko;
        private System.Windows.Forms.Button butSearchFileStanowisko;
        private System.Windows.Forms.Label pathStanowisko;
        private System.Windows.Forms.OpenFileDialog searchFileNaz;
        private System.Windows.Forms.OpenFileDialog searchFileSta;
        private System.Windows.Forms.Label createToFile;
        private System.Windows.Forms.Label pathToFile;
        private System.Windows.Forms.Button buttonSearchToFile;
        private System.Windows.Forms.OpenFileDialog searchToFile;
        private System.Windows.Forms.Button butCreateSheet;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label isOpenNazwisko;
        private System.Windows.Forms.Label isOpenToFile;
        private System.Windows.Forms.Label isOpenStan;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox passNazTurn;
        private System.Windows.Forms.TextBox passNazwisko;
        private System.Windows.Forms.CheckBox passStanTurn;
        private System.Windows.Forms.CheckBox passToFileTurn;
        private System.Windows.Forms.TextBox passStanowisco;
        private System.Windows.Forms.TextBox PassToFile;
    }
}

