namespace GMI_WS_Utilitario
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
            this.GMIUtilitario = new System.ServiceProcess.ServiceProcessInstaller();
            this.GMI_WS_Utilitario = new System.ServiceProcess.ServiceInstaller();
            // 
            // GMIUtilitario
            // 
            this.GMIUtilitario.Password = null;
            this.GMIUtilitario.Username = null;
            // 
            // GMI_WS_Utilitario
            // 
            this.GMI_WS_Utilitario.Description = "Utilitario para varias empresas, GRE Sunat";
            this.GMI_WS_Utilitario.DisplayName = "GRE_WS";
            this.GMI_WS_Utilitario.ServiceName = "Service1";
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.GMIUtilitario,
            this.GMI_WS_Utilitario});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller GMIUtilitario;
        private System.ServiceProcess.ServiceInstaller GMI_WS_Utilitario;
    }
}