namespace WS_GREUTILITARIO
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.serviceProcessInstallv9 = new System.ServiceProcess.ServiceProcessInstaller();
            this.serviceInstallerGrev9 = new System.ServiceProcess.ServiceInstaller();
            // 
            // serviceProcessInstallv9
            // 
            this.serviceProcessInstallv9.Password = null;
            this.serviceProcessInstallv9.Username = null;
            // 
            // serviceInstallerGrev9
            // 
            this.serviceInstallerGrev9.Description = "gre v9";
            this.serviceInstallerGrev9.DisplayName = "WS_GMI_v9";
            this.serviceInstallerGrev9.ServiceName = "Servicev9gre";
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.serviceProcessInstallv9,
            this.serviceInstallerGrev9});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller serviceProcessInstallv9;
        private System.ServiceProcess.ServiceInstaller serviceInstallerGrev9;
    }
}