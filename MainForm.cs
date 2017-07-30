using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace Duino_Controller_1._0
{
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        public MainForm()
        {
            InitializeComponent();
            this.StyleManager = metroStyleManager1;
            obterPortas();

        }
        // função que obtem as portas COM disponíveis para uso
        void obterPortas()
        {
            cmbPortas.Items.Clear();
            String[] portas = SerialPort.GetPortNames();
            cmbPortas.Items.AddRange(portas);
        }

        // Evento do Botão conectar
        private void butConectar_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbPortas.Text == "" || cmbBaud.Text == "")
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor selecione uma opção.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                }

                else
                {
                    serialPort.PortName = cmbPortas.Text;
                    serialPort.BaudRate = Convert.ToInt32(cmbBaud.Text);
                    try
                    {
                        serialPort.Open();
                        butTeste.Enabled = true;
                        butConectar.Enabled = false;
                        progressBarStatus.Value = 100;
                        lblConexoes.Text = "Conectado";
                    }

                    catch
                    {
                        MessageBox.Show("Porta não disponível", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
            }

            catch (UnauthorizedAccessException)
            {
                MessageBox.Show("Acesso Negado.");
            }
        }

        // Evento Botão desconectar
        private void bustDesconectar_Click(object sender, EventArgs e)
        {
            serialPort.Close();
            butTeste.Enabled = false;
            butConectar.Enabled = true;
            progressBarStatus.Value = 0;
            lblConexoes.Text = "Desconectado";
        }

        // Evento botão de Teste
        private void butTeste_Click(object sender, EventArgs e)
        {
            if (serialPort.IsOpen)
            {
                serialPort.Write("t");
                //string resposta = serialPort.ReadLine();
            }


        }
        // Método de escrever a resposta da serial do arduino no TextBox de teste
        public void escreverTxtTeste(object sender, EventArgs e)
        {
            txtTeste.Text = dataReceiverArduino;
        }

        public int valorLDR;
        // Atualiza a ProgressBar do LDR
        public void escreverTensao(object sender, EventArgs e)
        {
            //int valor = Convert.ToInt32(dataReceiverArduino);
            //double tensao = (valor * 5) / 1024;
            //lbTensao.Text = tensao.ToString() + " " + "v";

            ProgressBarLDR.SubscriptText = dataReceiverArduino;
            try
            {
                valorLDR = Convert.ToInt32(dataReceiverArduino);
            }

            catch
            {
                dataReceiverArduino = serialPort.ReadLine();
            }
            ProgressBarLDR.Value = valorLDR;
        }

        public double valorTensao;
        double tensaoRede;
        public void escreverTensaodaRede(object sender, EventArgs e)
        {
            try
            {
                valorTensao = Convert.ToDouble(tensaoReceiverArduino);
            }

            catch
            {
                tensaoReceiverArduino = serialPort.ReadLine();
            }

            tensaoRede = (valorTensao * 127) / 1024;
            ProgressBarTensaoREDE.Value = Convert.ToInt32(valorTensao);
            ProgressBarTensaoREDE.SubscriptText = tensaoRede.ToString("n2");
        }

        

        decimal temperaturaT;
        decimal umidadeD;
        public void escreveTemperatura(object sender, EventArgs e)
        {
            try
            {
                temperaturaT = Convert.ToDecimal(temperaturaReceiver);
                umidadeD = Convert.ToDecimal(umidadeReceiver);
            }

            catch
            {
                umidadeReceiver = serialPort.ReadLine();
            }


            ProgressBarTemperatura.Value = Convert.ToInt32(temperaturaT);
            ProgressBarTemperatura.SubscriptText = String.Format("{0:0:0}", temperaturaReceiver.ToString()); //temperatura.ToString("n2");
            if (umidadeReceiver != null)
            {

                ProgressBarUmidade.Value = Convert.ToInt32(umidadeD);
                ProgressBarUmidade.SubscriptText = String.Format("{0:0:0}", umidadeReceiver.Replace("\r", "").ToString()); //temperatura.ToString("n2");
            }
        }

        //Escreve Corrente
        decimal CorrenteC;
        string correnteReceiver;
        public void escreveCorrente(object sender, EventArgs e)
        {
            try
            {
                CorrenteC = Convert.ToDecimal(correnteReceiver);
                ProgressBarCorrente.Value = Convert.ToInt32(CorrenteC);
                ProgressBarCorrente.SubscriptText = String.Format("{0:0:0}", correnteReceiver.Replace("\r", "").ToString());
            }

            catch
            {

            }
           
        }

        // Evento de recebimento de dados pela porta serial 
        string dataReceiverArduino;
        public string tensaoReceiverArduino;
        string temperaturaReceiver;
        string umidadeReceiver;
        private void serialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            dataReceiverArduino = serialPort.ReadLine();
            dataReceiverArduino = dataReceiverArduino.Replace("\r", "");
            try
            {
                if (dataReceiverArduino == "t")
                {
                    dataReceiverArduino = serialPort.ReadLine();
                    this.Invoke(new EventHandler(escreverTxtTeste));
                }
            }

            catch
            {

            }

            try
            {
                if (dataReceiverArduino == "l")
                {
                    dataReceiverArduino = serialPort.ReadLine();
                    this.Invoke(new EventHandler(escreverTensao));
                }
            }
            catch
            {

            }


            try
            {
                if (dataReceiverArduino == "r")
                {
                    tensaoReceiverArduino = serialPort.ReadLine();
                    this.Invoke(new EventHandler(escreverTensaodaRede));
                }
            }

            catch
            {

            }

            try
            {
                if (dataReceiverArduino == "T")
                {
                    temperaturaReceiver = serialPort.ReadLine();

                    this.Invoke(new EventHandler(escreveTemperatura));

                }
            }

            catch
            {

            }

            try
            {
                if (dataReceiverArduino == "H")
                {

                    umidadeReceiver = serialPort.ReadLine();
                    this.Invoke(new EventHandler(escreveTemperatura));
                }
            }

            catch
            {

            }

            try
            {
                if (dataReceiverArduino == "P")
                {
                    if (statusSeguranca == true)
                    {
                        notifyIcon1.Text = "Presenaça detectada!";
                        notifyIcon1.BalloonTipTitle = "Alerta!!";
                        notifyIcon1.BalloonTipText = "Presença detectada!";
                        notifyIcon1.BalloonTipIcon = ToolTipIcon.Warning;
                        notifyIcon1.Visible = true;
                        notifyIcon1.ShowBalloonTip(3500);
                    }

                }
            }
            catch
            {

            }

            try
            {
                if (dataReceiverArduino == "a")
                {
                    correnteReceiver = serialPort.ReadLine();
                    this.Invoke(new EventHandler(escreveCorrente));
                }
            }

            catch
            {

            }

        }

        //Variáveis booleanas para controle dos relés pelos botões
        public bool rele1 = false;
        public bool rele2 = false;
        public bool rele3 = false;
        public bool rele4 = false;
        public bool rele5 = false;

        // botões que enviam os comandos para os relés
        private void tlRele1_Click(object sender, EventArgs e)
        {
            rele1 = !rele1;
            if (serialPort.IsOpen && rele1 == true)
            {
                serialPort.Write("1");
            }

            if (serialPort.IsOpen && rele1 == false)
            {
                serialPort.Write("A");
            }

            else if (serialPort.IsOpen == false)
            {
                MetroFramework.MetroMessageBox.Show(this, "Conecte a porta serial.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }

            if (lblStatusRele1.Text == "Desligado")
            {
                lblStatusRele1.Text = "Ligado";
            }
            else
            {
                lblStatusRele1.Text = "Desligado";
            }

        }

        private void tlRele2_Click(object sender, EventArgs e)
        {
            rele2 = !rele2;
            if (serialPort.IsOpen && rele2 == true)
            {
                serialPort.Write("2");
            }

            if (serialPort.IsOpen && rele2 == false)
            {
                serialPort.Write("B");
            }
            else if (serialPort.IsOpen == false)
            {
                MetroFramework.MetroMessageBox.Show(this, "Conecte a porta serial.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }

            if (lblStatusRele2.Text == "Desligado")
            {
                lblStatusRele2.Text = "Ligado";
            }
            else
            {
                lblStatusRele2.Text = "Desligado";
            }
        }

        private void tlRele3_Click(object sender, EventArgs e)
        {
            rele3 = !rele3;
            if (serialPort.IsOpen && rele3 == true)
            {
                serialPort.Write("3");
            }

            if (serialPort.IsOpen && rele3 == false)
            {
                serialPort.Write("C");
            }
            else if (serialPort.IsOpen == false)
            {
                MetroFramework.MetroMessageBox.Show(this, "Conecte a porta serial.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }

            if (lblStatusRele3.Text == "Desligado")
            {
                lblStatusRele3.Text = "Ligado";
            }
            else
            {
                lblStatusRele3.Text = "Desligado";
            }

        }

        private void tlRele4_Click(object sender, EventArgs e)
        {
            rele4 = !rele4;
            if (serialPort.IsOpen && rele4 == true)
            {
                serialPort.Write("4");
            }

            if (serialPort.IsOpen && rele4 == false)
            {
                serialPort.Write("D");
            }
            else if (serialPort.IsOpen == false)
            {
                MetroFramework.MetroMessageBox.Show(this, "Conecte a porta serial.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }

            if (lblStatusRele4.Text == "Desligado")
            {
                lblStatusRele4.Text = "Ligado";
            }
            else
            {
                lblStatusRele4.Text = "Desligado";
            }
        }

        private void tlRele5_Click(object sender, EventArgs e)
        {
            rele5 = !rele5;
            if (serialPort.IsOpen && rele5 == true)
            {
                serialPort.Write("5");
            }

            if (serialPort.IsOpen && rele5 == false)
            {
                serialPort.Write("E");
            }
            else if (serialPort.IsOpen == false)
            {
                MetroFramework.MetroMessageBox.Show(this, "Conecte a porta serial.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }
            if (lblStatusRele5.Text == "Desligado")
            {
                lblStatusRele5.Text = "Ligado";
            }
            else
            {
                lblStatusRele5.Text = "Desligado";
            }

        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------
        // configuração de temas e cores
        private void cmbTema_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            switch (cmbTema.SelectedIndex)
            {
                case 0:
                    metroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Dark;
                    ProgressBarLDR.BackColor = MetroFramework.MetroColors.Black;
                    ProgressBarTensaoREDE.BackColor = MetroFramework.MetroColors.Black;

                    //Salva as configurações do usuário
                    Properties.Settings.Default.thememetro = MetroFramework.MetroThemeStyle.Dark;
                    Properties.Settings.Default.themaprogres = "Black";
                    Properties.Settings.Default.Save();

                    break;
                case 1:
                    metroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Light;
                    ProgressBarLDR.BackColor = MetroFramework.MetroColors.White;
                    ProgressBarTensaoREDE.BackColor = MetroFramework.MetroColors.White;

                    //Salva as configurações do usuário
                    Properties.Settings.Default.thememetro = MetroFramework.MetroThemeStyle.Light;
                    Properties.Settings.Default.themaprogres = "White";
                    Properties.Settings.Default.Save();


                    break;
            }
        }

        private void cmbCor_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(cmbCor.SelectedIndex);
            //Salva as configurações do usuário
            Properties.Settings.Default.colorstyle = (MetroFramework.MetroColorStyle)Convert.ToInt32(cmbCor.SelectedIndex);
            Properties.Settings.Default.Save();
        }

        //Evento do botão Atualiza as portas
        private void butAtualizarPortas_Click(object sender, EventArgs e)
        {
            obterPortas();
        }

        // evento de controle da TrackbarSensibilidade
        public int valorTrackbarSensibilidade;
        private void TrackBarSensibilidade_ValueChanged(object sender, EventArgs e)
        {


            valorTrackbarSensibilidade = TrackBarSensibilidade.Value;
            lblValorTrackbar.Text = TrackBarSensibilidade.Value.ToString();
        }

        public bool liberaUso;
        //// Evento CheckBox Automático
        private void rdoUsarAutomatico_CheckedChanged(object sender, EventArgs e)
        {
            TrackBarSensibilidade.Maximum = 1023;
            TrackBarSensibilidade.Minimum = 0;
            TrackBarSensibilidade.Value = 950;
            TrackBarSensibilidade.Enabled = true;
            timerLDR.Enabled = true;
            liberaUso = true;

        }

        DateTime dataHora = DateTime.Now;
        // Propiedades carregadas na abertura do programa
        private void MainForm_Load(object sender, EventArgs e)
        {
            TrackBarSensibilidade.Enabled = false;

            lblConexoes.Text = "Desconectado";
            //listaValorLDRSource.DataSource = listaValorLDR;
            //dgvDados.DataSource = listaValorLDRSource;
            //dgvDados.ColumnCount = 10;
            //dgvDados.Columns[0].Name = "Tensão da Rede (V)";
            //dgvDados.Columns[1].Name = "Corrente (A)";
            //dgvDados.Columns[2].Name = "Temperatura";
            //dgvDados.Columns[3].Name = "Umidade (%)";
            //dgvDados.Columns[4].Name = "Taxa de Luminosidade (0 a 1023)";
            //dgvDados.Columns[5].Name = "Tempo";
            dgvDados.AutoResizeColumns();
            listaDadosSource.DataSource = listaDados;
            dgvDados.DataSource = listaDadosSource;



            //Lê as configuração de designer do usuário
            metroStyleManager1.Theme = Properties.Settings.Default.thememetro;

            if (Properties.Settings.Default.themaprogres == "Black")
            {

                ProgressBarTensaoREDE.BackColor = MetroFramework.MetroColors.Black;
                ProgressBarLDR.BackColor = MetroFramework.MetroColors.Black;
            }
            else if (Properties.Settings.Default.themaprogres == "White")
            {
                ProgressBarLDR.BackColor = MetroFramework.MetroColors.White;
                ProgressBarTensaoREDE.BackColor = MetroFramework.MetroColors.White;
            }
            metroStyleManager1.Style = Properties.Settings.Default.colorstyle;
            tlRele1.Text = Properties.Settings.Default.rele1;
            //rele1Enable = Properties.Settings.Default.rele1enable;
            //tlRele1.Enabled = !rele1Enable;
            //tlRele2.Text = Properties.Settings.Default.rele2;
            //rele2Enable = Properties.Settings.Default.rele2enable;
            //tlRele2.Enabled = !rele2Enable;
            //tlRele3.Text = Properties.Settings.Default.rele3;
            //rele3Enable = Properties.Settings.Default.rele3enable;
            //tlRele3.Enabled = !rele3Enable;
            //tlRele4.Text = Properties.Settings.Default.rele4;
            //rele4Enable = Properties.Settings.Default.rele4enable;
            //tlRele4.Enabled = !rele4Enable;
            //tlRele5.Text = Properties.Settings.Default.rele5;
            //rele5Enable = Properties.Settings.Default.rele5enable;
            //tlRele5.Enabled = !rele5Enable;
        }

        // Evento CheckBox Manual
        private void rdoManual_CheckedChanged(object sender, EventArgs e)
        {
            TrackBarSensibilidade.Enabled = false;
            liberaUso = false;
            timerLDR.Enabled = false;
            iluminacaoAutomatica();

        }

        //variaveis para Combobox RELE
        public bool rele1Enable;
        public bool rele2Enable;
        public bool rele3Enable;
        public bool rele4Enable;
        public bool rele5Enable;
        // Evento dos ComboBox de Configuração dos Botões
        private void cmbRele1_SelectedIndexChanged(object sender, EventArgs e)
        {

            switch (cmbRele1.SelectedIndex)
            {
                case 0:
                    tlRele1.Text = "Lâmpada";
                    tlRele1.Enabled = true;
                    rele1Enable = false;

                    break;

                case 1:
                    tlRele1.Text = "Ventilador";
                    tlRele1.Enabled = true;
                    rele1Enable = false;

                    break;

                case 2:
                    tlRele1.Text = "Motor";
                    tlRele1.Enabled = true;
                    rele1Enable = false;

                    break;
                case 3:
                    tlRele1.Text = "Iluminação Automática (LDR)";
                    tlRele1.Enabled = false;
                    rele1Enable = true;
                    break;


            }
            //Salva as configurações do usuário
            Properties.Settings.Default.rele1 = tlRele1.Text;
            Properties.Settings.Default.rele1enable = rele1Enable;
            Properties.Settings.Default.Save();
        }

        private void cmbRele2_SelectedIndexChanged(object sender, EventArgs e)
        {


            switch (cmbRele2.SelectedIndex)
            {
                case 0:
                    tlRele2.Text = "Lâmpada";
                    tlRele2.Enabled = true;
                    rele2Enable = false;

                    break;

                case 1:
                    tlRele2.Text = "Ventilador";
                    tlRele2.Enabled = true;
                    rele2Enable = false;

                    break;

                case 2:
                    tlRele2.Text = "Motor";
                    tlRele2.Enabled = true;
                    rele2Enable = false;

                    break;
                case 3:
                    tlRele2.Text = "Iluminação Automática (LDR)";
                    tlRele2.Enabled = false;
                    rele2Enable = true;

                    break;

            }
            //Salva as configurações do usuário
            Properties.Settings.Default.rele2 = tlRele2.Text;
            Properties.Settings.Default.rele2enable = rele2Enable;
            Properties.Settings.Default.Save();
        }

        private void cmbRele3_SelectedIndexChanged(object sender, EventArgs e)
        {

            switch (cmbRele3.SelectedIndex)
            {
                case 0:
                    tlRele3.Text = "Lâmpada";
                    tlRele3.Enabled = true;
                    rele3Enable = false;

                    break;

                case 1:
                    tlRele3.Text = "Ventilador";
                    tlRele3.Enabled = true;
                    rele3Enable = false;

                    break;

                case 2:
                    tlRele3.Text = "Motor";
                    tlRele3.Enabled = true;
                    rele3Enable = false;

                    break;
                case 3:
                    tlRele3.Text = "Iluminação Automática (LDR)";
                    tlRele3.Enabled = false;
                    rele3Enable = true;

                    break;

            }
            //Salva as configurações do usuário
            Properties.Settings.Default.rele3 = tlRele3.Text;
            Properties.Settings.Default.rele3enable = rele3Enable;
            Properties.Settings.Default.Save();
        }

        private void cmbRele4_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cmbRele4.SelectedIndex)
            {
                case 0:
                    tlRele4.Text = "Lâmpada";
                    tlRele4.Enabled = true;
                    rele4Enable = false;

                    break;

                case 1:
                    tlRele4.Text = "Ventilador";
                    tlRele4.Enabled = true;
                    rele4Enable = false;

                    break;

                case 2:
                    tlRele4.Text = "Motor";
                    tlRele4.Enabled = true;
                    rele4Enable = false;

                    break;
                case 3:
                    tlRele4.Text = "Iluminação Automática (LDR)";
                    tlRele4.Enabled = false;
                    rele4Enable = true;
                    break;

            }
            //Salva as configurações do usuário
            Properties.Settings.Default.rele4 = tlRele4.Text;
            Properties.Settings.Default.rele4enable = rele4Enable;
            Properties.Settings.Default.Save();
        }

        private void cmbRele5_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cmbRele5.SelectedIndex)
            {
                case 0:
                    tlRele5.Text = "Lâmpada";
                    tlRele5.Enabled = true;
                    rele5Enable = false;

                    break;

                case 1:
                    tlRele5.Text = "Ventilador";
                    tlRele5.Enabled = true;
                    rele5Enable = false;

                    break;

                case 2:
                    tlRele5.Text = "Motor";
                    tlRele5.Enabled = true;
                    rele5Enable = false;

                    break;
                case 3:
                    tlRele5.Text = "Iluminação Automática (LDR)";
                    tlRele5.Enabled = false;
                    rele5Enable = true;

                    break;

            }
            //Salva as configurações do usuário
            Properties.Settings.Default.rele5 = tlRele5.Text;
            Properties.Settings.Default.rele5enable = rele5Enable;
            Properties.Settings.Default.Save();
        }

        // Liga as lâmpadas de acordo com a configuração
        public void iluminacaoAutomatica()
        {

            if (rele1Enable == true && liberaUso == true && valorLDR >= valorTrackbarSensibilidade)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("1");
                }

            }

            if (rele1Enable == true && liberaUso == true && valorLDR <= valorTrackbarSensibilidade)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("A");
                }

            }

            if (rele2Enable == true && liberaUso && valorLDR >= valorTrackbarSensibilidade)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("2");
                }

            }

            if (rele2Enable == true && liberaUso == true && valorLDR <= valorTrackbarSensibilidade)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("B");
                }

            }

            if (rele3Enable == true && liberaUso == true && valorLDR >= valorTrackbarSensibilidade)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("3");
                }

            }

            if (rele3Enable == true && liberaUso == true && valorLDR <= valorTrackbarSensibilidade)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("C");
                }

            }

            if (rele4Enable == true && liberaUso == true && valorLDR >= valorTrackbarSensibilidade)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("4");
                }

            }

            if (rele4Enable == true && liberaUso == true && valorLDR <= valorTrackbarSensibilidade)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("D");
                }

            }

            if (rele5Enable == true && liberaUso == true && valorLDR >= valorTrackbarSensibilidade)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("5");
                }

            }

            if (rele5Enable == true && liberaUso == true && valorLDR <= valorTrackbarSensibilidade)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("E");
                }

            }

            if (rele1Enable == true && liberaUso == false)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("A");
                }

            }

            if (rele2Enable == true && liberaUso == false)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("B");
                }

            }

            if (rele3Enable == true && liberaUso == false)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("C");
                }

            }

            if (rele4Enable == true && liberaUso == false)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("D");
                }

            }

            if (rele5Enable == true && liberaUso == false)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write("E");
                }

            }


        }

        // Evento Timer de controle do Modo Iluminção automática LDR
        private void timerLDR_Tick(object sender, EventArgs e)
        {
            if (valorLDR >= valorTrackbarSensibilidade)
            {
                iluminacaoAutomatica();
            }

            if (valorLDR <= valorTrackbarSensibilidade)
            {
                iluminacaoAutomatica();
            }

           
        }

        // Quando o Programa é fechado executa essas funções
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (serialPort.IsOpen)
                serialPort.Close();
        }

        // cria objeto e lista para adicionar no DataGridView e Gráfico
        // Dados para LDR
        //LDR teste = new LDR();
        //IList<LDR> listaValorLDR = new BindingList<LDR>();
        //BindingSource listaValorLDRSource = new BindingSource();
        //// Dados para Tensão Da REDE
        //TensaoRede tensaoR = new TensaoRede();
        //IList<TensaoRede> listaValorTensaoRede = new BindingList<TensaoRede>();
        //BindingSource listaValorTensaoRedeSource = new BindingSource();
        dadosDatagridview dados = new dadosDatagridview();
        IList<dadosDatagridview> listaDados = new BindingList<dadosDatagridview>();
        BindingSource listaDadosSource = new BindingSource();
        public int acumulador;
        public void atualizarDataGridView()
        {
            dadosDatagridview dados = new dadosDatagridview();
            dados.tensaoRede = tensaoRede;
            dados.luminosidade = valorLDR;
            dados.temperatura = String.Format("{0:0:0}", temperaturaReceiver.ToString());
            dados.umidade = String.Format("{0:0:0}", umidadeReceiver.ToString());
            acumulador += timerMonitoramento.Interval;
            dados.corrente = correnteReceiver.ToString();
            dados.tempo += acumulador / 1000;
            listaDados.Add(dados);
            //if(dgvDados.Rows.Count > 15)
            //{
            //    dgvDados.FirstDisplayedScrollingRowIndex = dgvDados.Rows.Count;
            //}


        }


        // Evento para botão que reseta as configurações
        private void butreset_Click(object sender, EventArgs e)
        {
            //reseta os relés
            tlRele1.Text = "Relé 1";
            tlRele1.Enabled = true;
            tlRele2.Text = "Relé 2";
            tlRele2.Enabled = true;
            tlRele3.Text = "Relé 3";
            tlRele3.Enabled = true;
            tlRele4.Text = "Relé 4";
            tlRele4.Enabled = true;
            tlRele5.Text = "Relé 5";
            tlRele5.Enabled = true;
            Properties.Settings.Default.rele1 = tlRele1.Text;
            Properties.Settings.Default.rele1enable = false;
            Properties.Settings.Default.rele2 = tlRele2.Text;
            Properties.Settings.Default.rele2enable = false;
            Properties.Settings.Default.rele3 = tlRele3.Text;
            Properties.Settings.Default.rele3enable = false;
            Properties.Settings.Default.rele4 = tlRele4.Text;
            Properties.Settings.Default.rele4enable = false;
            Properties.Settings.Default.rele5 = tlRele5.Text;
            Properties.Settings.Default.rele5enable = false;
            metroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Light;
            ProgressBarLDR.BackColor = MetroFramework.MetroColors.White;
            ProgressBarTensaoREDE.BackColor = MetroFramework.MetroColors.White;
            Properties.Settings.Default.thememetro = MetroFramework.MetroThemeStyle.Light;
            Properties.Settings.Default.themaprogres = MetroFramework.MetroColors.White.ToString();
            metroStyleManager1.Style = MetroFramework.MetroColorStyle.Default;
            Properties.Settings.Default.colorstyle = MetroFramework.MetroColorStyle.Default;
            Properties.Settings.Default.Save();

        }

        // seleciona o tempo de atualização dos dados 
        int tempoParaTimerAtualizacao = 0;
        private void cmbTaxaAtualizacao_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cmbTaxaAtualizacao.SelectedIndex)
            {
                case 0:
                    tempoParaTimerAtualizacao = 1000;
                    break;

                case 1:
                    tempoParaTimerAtualizacao = 5000;
                    break;

                case 2:
                    tempoParaTimerAtualizacao = 10000;
                    break;

                case 3:
                    tempoParaTimerAtualizacao = 30000;
                    break;

                case 4:
                    tempoParaTimerAtualizacao = 60000;
                    break;

                case 5:
                    tempoParaTimerAtualizacao = 300000;
                    break;
                case 7:
                    tempoParaTimerAtualizacao = 18000000;
                    break;
            }
        }

        // Evento para botão de iniciar o monitoramento
        private void butIniciarMonitoramento_Click(object sender, EventArgs e)
        {
            if (serialPort.IsOpen)
            {
                if (tempoParaTimerAtualizacao > 0)
                {
                    timerMonitoramento.Interval = tempoParaTimerAtualizacao;
                    timerMonitoramento.Enabled = true;
                    butIniciarMonitoramento.Enabled = false;
                    cmbTaxaAtualizacao.Enabled = false;
                }

                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Selecione um tempo de atualização.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

            }

            else
            {
                MetroFramework.MetroMessageBox.Show(this, "A porta serial está desconectada!", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }
        // Estouro do timer de monitoramento
        private void timerMonitoramento_Tick(object sender, EventArgs e)
        {
            atualizarDataGridView();
        }

        // Evento de parar atualização 
        private void butPararMonitoramento_Click(object sender, EventArgs e)
        {
            timerMonitoramento.Enabled = false;
            butIniciarMonitoramento.Enabled = true;
            cmbTaxaAtualizacao.Enabled = true;
            acumulador = 0;
        }

        //Evento do Botão apagar dados
        private void butApagarDados_Click(object sender, EventArgs e)
        {
            dgvDados.Rows.Clear();
        }

        //BackGroundWorker Exportar para Excel
        private void backgroundWorkerExportar_DoWork(object sender, DoWorkEventArgs e)
        {
            string filename = ((DataParametros)e.Argument).fileName;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wbExcel = excel.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet wsExcel = (Worksheet)excel.ActiveSheet;
            excel.Visible = false;
            int index = 1;
            int listaQuantidade = listaDados.Count;
            wsExcel.Cells[1, 1] = "Tensão da Rede (V)";
            wsExcel.Cells[1, 2] = "Corrente (A)";
            wsExcel.Cells[1, 3] = "Temperatura (°C)";
            wsExcel.Cells[1, 4] = "Umidade (%)";
            wsExcel.Cells[1, 5] = "Luminosidade (0 a 1023)";
            wsExcel.Cells[1, 6] = "Tempo (s)";

            foreach (dadosDatagridview d in listaDados)
            {
                if (!backgroundWorkerExportar.CancellationPending)
                {
                    backgroundWorkerExportar.ReportProgress(index++ * 100 / listaQuantidade);
                    wsExcel.Cells[index, 1] = d.tensaoRede;
                    wsExcel.Cells[index, 2] = d.corrente;
                    wsExcel.Cells[index, 3] = d.temperatura;
                    wsExcel.Cells[index, 4] = d.umidade;
                    wsExcel.Cells[index, 5] = d.luminosidade;
                    wsExcel.Cells[index, 6] = d.tempo;
                }
            }

            wsExcel.SaveAs(filename, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges,
                Type.Missing, Type.Missing);
            excel.Quit();
        }

        private void backgroundWorkerExportar_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBarExportar.Value = e.ProgressPercentage;
            ProgressBarExportar.Update();
        }

        private void backgroundWorkerExportar_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                Thread.Sleep(100);
                MetroFramework.MetroMessageBox.Show(this, "Exportado com sucesso!", "Mensagem", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
        }

        struct DataParametros
        {

            public string fileName { get; set; }
        }

        DataParametros entradaParametros;

        private void butExportar_Click(object sender, EventArgs e)
        {
            if (backgroundWorkerExportar.IsBusy)
                return;
            using (SaveFileDialog exporExcel = new SaveFileDialog() { Filter = "Excel |*.xls" })
            {
                if (exporExcel.ShowDialog() == DialogResult.OK)
                {
                    entradaParametros.fileName = exporExcel.FileName;
                    ProgressBarExportar.Minimum = 0;
                    ProgressBarExportar.Value = 0;
                    backgroundWorkerExportar.RunWorkerAsync(entradaParametros);
                }
            }
        }

        private void timerTemperatura_Tick(object sender, EventArgs e)
        {
            if (serialPort.IsOpen)
            {
                serialPort.Write("T");
            }

        }

        //radiobox ligar de deligar seguraça
        bool statusSeguranca;
        private void rdoSegurancaLigado_CheckedChanged(object sender, EventArgs e)
        {
            statusSeguranca = true;
        }

        private void rdoSegurancaDesligado_CheckedChanged(object sender, EventArgs e)
        {
            statusSeguranca = false;
        }

    }    
}
