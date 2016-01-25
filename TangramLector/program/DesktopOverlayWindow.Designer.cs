namespace tud.mci.tangram.TangramLector
{
    partial class DesktopOverlayWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.totalDisplayTimer = new System.Windows.Forms.Timer(this.components);
            this.onOffStepTimer = new System.Windows.Forms.Timer(this.components);
            this.transparentOverlayPictureBox1 = new tud.mci.tangram.TangramLector.TransparentOverlayPictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.transparentOverlayPictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // totalDisplayTimer
            // 
            this.totalDisplayTimer.Tick += new System.EventHandler(this.timer_Tick);
            // 
            // onOffStepTimer
            // 
            this.onOffStepTimer.Tick += new System.EventHandler(this.onOffStepTimer_Tick);
            // 
            // transparentOverlayPictureBox1
            // 
            this.transparentOverlayPictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.transparentOverlayPictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.transparentOverlayPictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.transparentOverlayPictureBox1.Location = new System.Drawing.Point(0, 0);
            this.transparentOverlayPictureBox1.Margin = new System.Windows.Forms.Padding(0);
            this.transparentOverlayPictureBox1.Name = "transparentOverlayPictureBox1";
            this.transparentOverlayPictureBox1.Size = new System.Drawing.Size(341, 283);
            this.transparentOverlayPictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.transparentOverlayPictureBox1.TabIndex = 0;
            this.transparentOverlayPictureBox1.TabStop = false;
            // 
            // DesktopOverlayWindow
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.Red;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(341, 283);
            this.ControlBox = false;
            this.Controls.Add(this.transparentOverlayPictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "DesktopOverlayWindow";
            this.Opacity = 0.5D;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.TopMost = true;
            this.TransparencyKey = System.Drawing.Color.Red;
            ((System.ComponentModel.ISupportInitialize)(this.transparentOverlayPictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private TransparentOverlayPictureBox transparentOverlayPictureBox1;
        private System.Windows.Forms.Timer totalDisplayTimer;
        private System.Windows.Forms.Timer onOffStepTimer;





    }
}