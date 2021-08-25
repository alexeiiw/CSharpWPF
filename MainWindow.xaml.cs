using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ScoreInterno
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        // Define lista de valores
        string[] lstMarcas = {"Yamaha","HP","New Holland","Massey Ferguson","Hyundai","Vesa","Isuzu","Canon"};
        string[] lstTiendas = {"Yamaha Central", "Yamaha Zona 18", "Yamaha Portales", "Yamaha El Punto", "Yamaha Roosevelt", "Yamaha Flagship", "Yamaha El Frutal", "Yamaha Villa Nueva", 
            "Yamaha Amatitlan", "Vesa Central", "Hp Central", "Home Office Majadas", "Home Office Oakland", "Isuzu Central", "Distribuidor Petapa", "Distribudidor Cs de la Cruz", 
            "Distribuidor Difisa", "Distribuidor Motowork", "Distribuidor Vegua", "Maquinaria Central"};
        //string[] lstVendedores = {"Axel Ceballos", "Bagner Bamaca", "Benjamin Turcios", "Boris Romero", "Charly Reyes", "Crishna Lopez", "Cristian Villar", "Daniel Balcarcel",
        //    "Daniel Barcarcel", "Daniel Pardo", "Danilo Quiñonez", "Denis Corado", "Denis Illescas", "Diego Cardona", "Diego Contreras", "Edilcar Quiñonez", "Eduin Ortega", "Emily de Leon",
        //    "Engelber Ramirez", "Evelyn Gutierrez", "Felyx", "Francisco ", "Fredy Cumez", "Glenda Mejia", "Henry Arce", "Henry Florian", "Hugo Mixteco", "Hugo Siguenza", "Jeimmy Chan",
        //    "Jimy Castillo", "Jorge Bautista", "Jorge Garcia", "Jorge Vasquez", "Jorlenny Gudiel", "Jose Miguel Garcia", "Josselyn Castillo", "Juan Carlos Gonzalez", "Kevin Ibañez",
        //    "Laura Reyes", "Laura Tohom", "Marco Tulio Villacinda", "Mariana Castillo", "Mario Avea", "Mauro Sanchez", "Merlin Barahona", "Mishel Calderon", "Oliver Vasquez", "Oscar Cuellar",
        //    "Osiris Chavez", "Osiris Chavez", "Ozman Lol", "Piter Contreras", "Roberto Godinez", "Rodrigo Muy", "Ruben Arana", "Ruben Reyes", "Seever Olivares", "Selvin Rac", "Sergio Lopez",
        //    "Sergio Moreno", "Sergio Roques", "Steve Taylor", "Viviana Quiroa", "William Perez", "Yeimi Garcia"};
        string[] lstTipoClientes = { "Asalariado Formal", "Asalariado Informal", "Comerciante Formal", "Comerciante Informal" };
        string[] lstTipoViviendas = { "Alquilada", "Propia", "Familiar" };
        string[] lstTipoCreditos = { "0 % Enganche", "10 % Enganche", "35 % Enganche", "Cliente Recurrente", "Credito a Colaborador", "Plan 0-24" };
        string[] lstRelaciones = { "Compañero de Trabajo", "Amigo", "Vecino", "Novio/a", "Cuñado/a" };
        string[] lstParentescos = { "Primo/a", "Tio/a", "Mamá", "Papá", "Padrastro/Madrastra", "Hermano/a", "Sobrino/a" };
        string[] lstDepartamentos = { "Alta Verapaz", "Baja Verapaz", "Chimaltenago", "Chiquimula", "Guatemala", "El Progreso", "Escuintla", "Huehuetenango", "Izabal", "Jalapa", "Jutiapa", 
            "Petén", "Quetzaltenango", "Quiché", "Retalhuleu", "Sacatepequez", "San Marcos", "Santa Rosa", "Sololá", "Suchitepequez", "Totonicapán", "Zacapa" };
        string[] lstAcciones = { "Aprobar", "Regresar", "Rechazar" };

        // Define objetos globales
        //string connStringUTILS = "data source=128.1.200.169;initial catalog=UTILS_SCR;persist security info=True;user id=usrsap;password=C@nella20$"; // Producción                                                                                                                                                      // Define objetos globales
        string connStringUTILS = "data source=128.1.202.136;initial catalog=UTILS;persist security info=True;user id=usrsap;password=C@nella20$"; // Desarrollo

        string connStringSTOD = "data source=128.1.200.180;initial catalog=STOD_SAPBONE;persist security info=True;user id=usrSTOD;password=stod20$";

        string Sistema = "";

        string Rol = "";

        string Vendedor = "";

        string SerieAPP = "gV2xPUmL";

        public MainWindow()
        {
            InitializeComponent();
        }

        // Evento cuando se inicia el formulario
        private void OnLoad(object sender, RoutedEventArgs e)
        {
            // Valida que el número de serie de la aplicación sea válido
            if (Validar_SerieAPP()) { Iniciar_Aplicacion(); }
            else
            {
                MessageBox.Show("Su versión de aplicación ya no es válida, favor comunicarse con el Administrador!", "SCORE INTERNO", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Windows.Application.Current.Shutdown();
            }
        }

        // Evento del boton accion
        private void btnAccion_Click(object sender, RoutedEventArgs e)
        {
            switch(btnAccion.Content)
            {
                case "ENVIAR A CREDITOS":
                    var objResultado1 = MessageBox.Show("Desea enviar la solicitud a créditos?", Sistema, MessageBoxButton.YesNo, MessageBoxImage.Question);
                    string strError1 = "";

                    if (objResultado1 == MessageBoxResult.Yes)
                    {
                        if (Guardar_Estado(Convert.ToInt32(txtSolicitud.Text), 2, ref strError1))
                        {
                            MessageBox.Show("La solicitud ha sido enviada a créditos satisfactoriamente!", Sistema, MessageBoxButton.OK, MessageBoxImage.Information);

                            Limpiar_Controles();

                            Bloquear_Controles();

                            Obtener_Solicitudes();

                            Aplicar_PermisosRol();

                            tabSolicitudes.Focus();

                            btnABC.Content = "Nuevo";
                        }
                        else { MessageBox.Show("La solicitud no fue enviada a créditos! " + "Error: " + strError1, Sistema, MessageBoxButton.OK, MessageBoxImage.Error); }
                    }

                    break;

                case "ACTUALIZAR":
                    var objResultado2 = MessageBox.Show("Desea "+cmbAccion.SelectedItem.ToString()+" la solicitud?", Sistema, MessageBoxButton.YesNo, MessageBoxImage.Question);
                    string strError2 = "";

                    int intEstado = 0;
                    switch(cmbAccion.SelectedItem.ToString())
                    {
                        case "APROBAR":
                            intEstado = 99;
                            break;
                        case "REGRESAR":
                            intEstado = 98;
                            break;
                        case "RECHAZAR":
                            intEstado = 97;
                            break;
                        default:
                            break;
                    }

                    if (objResultado2 == MessageBoxResult.Yes)
                    {
                        if (Guardar_Estado(Convert.ToInt32(txtSolicitud.Text), intEstado, ref strError2))
                        {
                            MessageBox.Show("La solicitud ha sido actualizada satisfactoriamente!", Sistema, MessageBoxButton.OK, MessageBoxImage.Information);

                            Limpiar_Controles();

                            Bloquear_Controles();

                            Obtener_Solicitudes();

                            Aplicar_PermisosRol();

                            tabSolicitudes.Focus();
                        }
                        else { MessageBox.Show("La solicitud no fue actualizada! " + "Error: " + strError2, Sistema, MessageBoxButton.OK, MessageBoxImage.Error); }
                    }

                    break;
            }
        }

        // Evento de cancelar 
        private void btnCancelar_Click(object sender, RoutedEventArgs e)
        {
            // Bloquear la aplicación
            Limpiar_Controles();

            Bloquear_Controles();

            Aplicar_PermisosRol();

            tabSolicitudes.Focus();
            tabSolicitudes.Focus();
        }

        // Evento para calcular el total de los ingresos
        private void Total_IngresosCliente(object sender, EventArgs e)
        {
            txtTotalIngresos.IsEnabled = false;

            if (txtIngresosFijos.Text == "") { txtIngresosFijos.Text = "0.00"; }
            if (txtIngresosVariables.Text == "") { txtIngresosVariables.Text = "0.00"; }
            if (txtOtrosIngresos.Text == "") { txtOtrosIngresos.Text = "0.00"; }

            txtTotalIngresos.Text = Aplicar_FormatoMoneda((Convert.ToDouble(txtIngresosFijos.Text) + Convert.ToDouble(txtIngresosVariables.Text) + Convert.ToDouble(txtOtrosIngresos.Text)), 1);
        }

        // Evento que valida que el textbox no este vacio
        private void TextBox_Obligatorio(object sender, EventArgs e)
        {
            TextBox objTextbox = sender as TextBox; 
            if (objTextbox.Text == "") 
            {
                // Habilitar para hacer las pruebas de datos obligatorios
                MessageBox.Show("El dato es obligatorio!", Sistema, MessageBoxButton.OK, MessageBoxImage.Error);
                // Regresa el foco al textbox
                //objTextbox.Dispatcher.BeginInvoke((Action)(() => { objTextbox.Focus(); }));
            } 
        }

        // Evento que valida el dpi
        private void Validar_DPI(object sender, EventArgs e)
        {
            TextBox objTextbox = sender as TextBox;
            if (objTextbox.Text.Count() != 13)
            {
                MessageBox.Show("El DPI no es valido!", Sistema, MessageBoxButton.OK, MessageBoxImage.Error);
                // Regresa el foco al textbox
                objTextbox.Dispatcher.BeginInvoke((Action)(() => { objTextbox.Focus(); }));
            }
        }

        // Evento que valida el teléfono
        private void Validar_Telefono(object sender, EventArgs e)
        {
            TextBox objTextbox = sender as TextBox;
            if (objTextbox.Text.Count() != 8)
            {
                MessageBox.Show("El teléfono no es valido!", Sistema, MessageBoxButton.OK, MessageBoxImage.Error);
                // Regresa el foco al textbox
                //objTextbox.Dispatcher.BeginInvoke((Action)(() => { objTextbox.Focus(); }));
            }
        }

        // Evento que valida el correo
        private void Validar_CorreoElectronico(object sender, EventArgs e)
        {
            TextBox objTextbox = sender as TextBox;
            if (Validar_CorreoElectronico(objTextbox.Text) == false)
            {
                MessageBox.Show("El correo electrónico no es valido!", Sistema, MessageBoxButton.OK, MessageBoxImage.Error);
                // Regresa el foco al textbox
                objTextbox.Dispatcher.BeginInvoke((Action)(() => { objTextbox.Focus(); }));
            }
        }

        // Evento que valida la fecha
        private void Validar_Fecha(object sender, EventArgs e)
        {
            TextBox objTextbox = sender as TextBox;
            if (Validar_Fecha(objTextbox.Text) == false)
            {
                MessageBox.Show("El formato de la fecha no es valido!", Sistema, MessageBoxButton.OK, MessageBoxImage.Error);
                // Regresa el foco al textbox
                objTextbox.Dispatcher.BeginInvoke((Action)(() => { objTextbox.Focus(); }));
            }
        }

        // Acción del boton acción
        private void btnDetalleAccion_Click(object sender, RoutedEventArgs e)
        {
            Button btnBoton = sender as Button;

            Llenar_Solicitud(Convert.ToInt32(btnBoton.Tag));

            Obtener_Comentarios();

            btnABC.Content = "Nuevo";
            btnABC.IsEnabled = true;
            lblMensajes.Visibility = Visibility.Hidden;
            txtDatosSolicitud.Visibility = Visibility.Hidden;

            switch (txtEstadoSolicitud.Text)
            {
                case "Enviado a Créditos":
                    string strUsuario = "";
                    string strComentarios = "";
                    string strFecha = "";
                    int intEstado = 0;

                    Obtener_DatosEstadoSolicitud(ref intEstado, ref strFecha, ref strUsuario, ref strComentarios);

                    txtDatosSolicitud.Text = Convertir_Fecha(strFecha, 2);

                    break;

                case "Ingresado":
                    btnABC.Content = "Habilitar";

                    if (Validar_EstadoCreditos())
                    {
                        btnAccion.Content = "ENVIAR A CREDITOS";
                        btnAccion.IsEnabled = true;
                        txtComentarios.IsEnabled = true;
                    }
                    else
                    {
                        lblMensajes.Visibility = Visibility.Visible;
                        lblMensajes.Content = "Hace falta llenar datos obligatorios para enviar la solicitud a Créditos!";
                    }

                    break;

                case "Aprobado":

                    break;

                case "Rechazado":

                    break;

                case "Retornado":
                    btnABC.Content = "Habilitar";

                    if (Validar_EstadoCreditos())
                    {
                        btnAccion.Content = "ENVIAR A CREDITOS";
                        btnAccion.IsEnabled = true;
                        txtComentarios.IsEnabled = true;
                    }
                    else
                    {
                        lblMensajes.Visibility = Visibility.Visible;
                        lblMensajes.Content = "Hace falta llenar datos obligatorios para enviar la solicitud a Créditos!";
                    }

                    break;

                default:
                    break;
            }

            tabSolicitud.Focus();
        }

        // Valida que existan todos los datos obligatorios para enviar la solicitud a creditos
        private bool Validar_EstadoCreditos()
        {
            if (txtTelefonoVendedor.Text == "") { return false; }
            if (txtPrimerNombreCliente.Text == "") { return false; }
            if (txtPrimerApellidoCliente.Text == "") { return false; }
            if (txtDpi.Text == "") { return false; }
            if (txtNit.Text == "") { return false; }
            if (txtFechaNacimiento.Text == "") { return false; }
            if (txtEstadoCivil.Text == "") { return false; }
            if (txtSexo.Text == "") { return false; }
            if (txtDependientes.Text == "") { return false; }
            if (txtProfesion.Text == "") { return false; }
            if (txtDireccion.Text == "") { return false; }
            if (txtColonia.Text == "") { return false; }
            if (txtMunicipio.Text == "") { return false; }
            if (txtZona.Text == "") { return false; }
            if (txtTelefonoCasa.Text == "") { return false; }
            if (txtTelefonoMovil.Text == "") { return false; }
            if (txtCorreoElectronico.Text == "") { return false; }
            if (txtNombreEmpresaCliente.Text == "") { return false; }
            if (txtCargo.Text == "") { return false; }
            if (txtFechaIngreso.Text == "") { return false; }
            if (txtIngresosFijos.Text == "") { return false; }
            if (txtDireccionTrabajoCliente.Text == "") { return false; }
            if (txtColoniaTrabajoCliente.Text == "") { return false; }
            if (txtMunicipioTrabajoCliente.Text == "") { return false; }
            if (txtZonaTrabajoCliente.Text == "") { return false; }
            if (txtJefeInmediato.Text == "") { return false; }
            if (txtTelefonoTrabajo.Text == "") { return false; }
            if (txtTelefonoJefe.Text == "") { return false; }
            if (txtNitEmpresa.Text == "") { return false; }
            if (txtTotalIngresos.Text == "") { return false; }
            if (txtNombreReferenciaPersonal1.Text == "") { return false; }
            if (txtTelefonoMovilReferenciaPersonal1.Text == "") { return false; }
            if (txtNombreReferenciaPersonal2.Text == "") { return false; }
            if (txtTelefonoMovilReferenciaPersonal2.Text == "") { return false; }
            if (txtNombreReferenciaPersonal3.Text == "") { return false; }
            if (txtTelefonoMovilReferenciaPersonal3.Text == "") { return false; }
            if (txtNombreReferenciaFamiliar1.Text == "") { return false; }
            if (txtTelefonoMovilReferenciaFamiliar1.Text == "") { return false; }
            if (txtNombreReferenciaFamiliar2.Text == "") { return false; }
            if (txtTelefonoMovilReferenciaFamiliar2.Text == "") { return false; }
            if (txtNombreReferenciaFamiliar3.Text == "") { return false; }
            if (txtTelefonoMovilReferenciaFamiliar3.Text == "") { return false; }
            if (txtArticulo.Text == "") { return false; }
            if (txtPrecio.Text == "") { return false; }
            if (txtPlazo.Text == "") { return false; }
            if (txtEnganche.Text == "") { return false; }
            if (txtValorCuota.Text == "") { return false; }
            if (txtSaldo.Text == "") { return false; }

            if (txtEstadoSolicitud.Text == "Enviado a Créditos") { return false; }

            return true; 
        }

        // Elabora el formato para la moneda
        private string Aplicar_FormatoMoneda(double dblValor, byte bytTipo)
        {
            if (bytTipo == 1)
            {
                if (dblValor == 0.00) { return "0"; }
                else { return dblValor.ToString("Q###,###,###.00"); }
            }
            if (bytTipo == 2)
            {
                if (dblValor == 0.00) { return "0.00"; }
                else { return dblValor.ToString("###,###,###.00"); }
            }
            else { return ""; }
        }

        // Metodo que valida la serie y la vigencia de la aplicacion
        private bool Validar_SerieAPP()
        {
            string strSql = "select * from [128.1.200.136].UTILS.dbo.SeriesAplicaciones where Serie = '" + SerieAPP + "'";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUTILS))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(strSql, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            if (dt.Rows.Count > 0)
            {
                if (bool.Parse(dt.Rows[0]["Activo"].ToString()))
                {
                    Title = dt.Rows[0]["Aplicacion"].ToString() + " Versión: " + dt.Rows[0]["Version"].ToString() + " " + dt.Rows[0]["Ambiente"].ToString();
                    Sistema = dt.Rows[0]["Aplicacion"].ToString();
                    return true;
                }
                else { return false; }
            }
            else { return false; }
        }

        // Llenar la solicitud
        private void Llenar_Solicitud(int intSolicitud)
        {
            string strSql = @"select (select top 1 ESTADO from SI_Estados where ID_SOLICITUD = SO.ID order by id desc) as estado, SO.* 
                            from SI_Solicitudes SO where SO.ID = " + intSolicitud;

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUTILS))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(strSql, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            // Dibujar el detalle
            if (dt.Rows.Count > 0)
            {
                txtSolicitud.Text = dt.Rows[0]["id"].ToString();
                switch (dt.Rows[0]["estado"].ToString()) 
                {
                    case "1":
                        txtEstadoSolicitud.Text = "Ingresado";
                        break;
                    case "2":
                        txtEstadoSolicitud.Text = "Enviado a Créditos";
                        break;
                    case "99":
                        txtEstadoSolicitud.Text = "Aprobado";
                        break;
                    case "98":
                        txtEstadoSolicitud.Text = "Retornado";
                        break;
                    case "97":
                        txtEstadoSolicitud.Text = "Rechazado";
                        break;
                    default:
                        txtEstadoSolicitud.Text = "";
                        break;
                }
                cmbMarca.SelectedItem = dt.Rows[0]["marca"].ToString();
                cmbTienda.SelectedItem = dt.Rows[0]["tienda"].ToString();
                txtTelefonoVendedor.Text = dt.Rows[0]["telefono_vendedor"].ToString();
                cmbTipoCliente.SelectedItem = dt.Rows[0]["tipo_cliente"].ToString();
                txtFechaSolicitud.Text = Convertir_Fecha(dt.Rows[0]["fecha_solicitud"].ToString(), 2);
                cmbVendedor.Items.Insert(1, dt.Rows[0]["nombre_vendedor"].ToString());
                cmbVendedor.SelectedIndex = 1;

                txtPrimerNombreCliente.Text = dt.Rows[0]["PRIMER_NOMBRE"].ToString();
                txtSegundoNombreCliente.Text = dt.Rows[0]["SEGUNDO_NOMBRE"].ToString();
                txtPrimerApellidoCliente.Text = dt.Rows[0]["PRIMER_APELLIDO"].ToString();
                txtSegundoApellidoCliente.Text = dt.Rows[0]["SEGUNDO_APELLIDO"].ToString();
                txtApellidoCasadaCliente.Text = dt.Rows[0]["APELLIDO_CASADA"].ToString();
                txtDpi.Text = dt.Rows[0]["DPI_PASAPORTE_CLIENTE"].ToString();
                txtNacionalidad.Text = dt.Rows[0]["NACIONALIDAD"].ToString();
                txtNit.Text = dt.Rows[0]["NIT_CLIENTE"].ToString();
                txtFechaNacimiento.Text = Convertir_Fecha(dt.Rows[0]["FECHA_NACIMIENTO_CLIENTE"].ToString(),2);
                txtEstadoCivil.Text = dt.Rows[0]["ESTADO_CIVIL"].ToString();
                txtSexo.Text = dt.Rows[0]["SEXO"].ToString();
                txtDependientes.Text = dt.Rows[0]["DEPENDIENTE"].ToString();
                txtProfesion.Text = dt.Rows[0]["PROFESION"].ToString();
                txtDireccion.Text = dt.Rows[0]["DIRECCION_CLIENTE"].ToString();
                txtColonia.Text = dt.Rows[0]["COLONIA_CLIENTE"].ToString();
                cmbDepartamento.SelectedItem = dt.Rows[0]["DEPARTAMENTO_CLIENTE"].ToString();
                txtMunicipio.Text = dt.Rows[0]["MUNICIPIO_CLIENTE"].ToString();
                txtZona.Text = dt.Rows[0]["ZONA_CLIENTE"].ToString();
                cmbTipoVivienda.SelectedItem = dt.Rows[0]["TIPO_VIVIENDA"].ToString();
                txtTelefonoCasa.Text = dt.Rows[0]["TELEFONO_CASA"].ToString();
                txtTelefonoMovil.Text = dt.Rows[0]["TELEFONO_MOVIL"].ToString();
                txtCorreoElectronico.Text = dt.Rows[0]["CORREO_ELECTRONICO"].ToString();

                txtNombreEmpresaCliente.Text = dt.Rows[0]["NOMBRE_EMPRESA"].ToString();
                txtCargo.Text = dt.Rows[0]["CARGO_CLIENTE"].ToString();
                txtFechaIngreso.Text = Convertir_Fecha(dt.Rows[0]["FECHA_INGRESO_CLIENTE"].ToString(),2);
                txtIngresosFijos.Text = Aplicar_FormatoMoneda(Convert.ToDouble(dt.Rows[0]["INGRESOS_FIJOS"]), 1);
                txtIngresosVariables.Text = Aplicar_FormatoMoneda(Convert.ToDouble(dt.Rows[0]["INGRESOS_VARIABLES"]), 1);
                txtOtrosIngresos.Text = Aplicar_FormatoMoneda(Convert.ToDouble(dt.Rows[0]["INGRESOS_OTROS"]), 1);
                txtFuenteOtrosIngresos.Text = dt.Rows[0]["FUENTES_INGRESOS_OTROS"].ToString();
                txtDireccionTrabajoCliente.Text = dt.Rows[0]["DIRECCION_EMPRESA"].ToString();
                txtColoniaTrabajoCliente.Text = dt.Rows[0]["COLONIA_EMPRESA"].ToString();
                cmbDepartamentoTrabajoCliente.SelectedItem = dt.Rows[0]["DEPARTAMENTO_EMPRESA"].ToString();
                txtMunicipioTrabajoCliente.Text = dt.Rows[0]["MUNICIPIO_EMPRESA"].ToString();
                txtZonaTrabajoCliente.Text = dt.Rows[0]["ZONA_EMPRESA"].ToString();
                txtJefeInmediato.Text = dt.Rows[0]["JEFE_INMEDIATO"].ToString();
                txtTelefonoTrabajo.Text = dt.Rows[0]["TELEFONO_TRABAJO"].ToString();
                txtTelefonoJefe.Text = dt.Rows[0]["TELEFONO_OFICINA"].ToString();
                txtExtension.Text = dt.Rows[0]["TELEFONO_EXTENSION"].ToString();
                txtNitEmpresa.Text = dt.Rows[0]["NIT_EMPRESA"].ToString();

                double dblIngresosFijos = 0;
                double dblIngresosVariables = 0;
                double dblOtrosIngresos = 0;
                if (txtIngresosFijos.Text == "") { dblIngresosFijos = 0; }
                else { dblIngresosFijos = Convert.ToDouble(txtIngresosFijos.Text.Replace("Q", "")); }
                if (txtIngresosVariables.Text == "") { dblIngresosVariables = 0; }
                else { dblIngresosVariables = Convert.ToDouble(txtIngresosVariables.Text.Replace("Q", ""));  }
                if (txtOtrosIngresos.Text == "") { dblOtrosIngresos = 0; }
                else { dblOtrosIngresos = Convert.ToDouble(txtOtrosIngresos.Text.Replace("Q", "")); }
                txtTotalIngresos.Text = Aplicar_FormatoMoneda((dblIngresosFijos + dblIngresosVariables + dblOtrosIngresos), 1);

                txtNombresApellidosConyuge.Text = dt.Rows[0]["NOMBRES_APELLIDOS_CONYUGE"].ToString();
                txtDpiConyuge.Text = dt.Rows[0]["DPI_PASAPORTE_CONYUGE"].ToString();
                txtFechaNacimientoConyuge.Text = Convertir_Fecha(dt.Rows[0]["FECHA_NACIMIENTO_CONYUGE"].ToString(),2);
                txtTrabajoConyuge.Text = dt.Rows[0]["TRABAJO_CONYUGE"].ToString();
                txtCargoConyuge.Text = dt.Rows[0]["CARGO_CONYUGE"].ToString();
                txtFechaIngresoConyuge.Text = Convertir_Fecha(dt.Rows[0]["FECHA_INGRESO_CONYUGE"].ToString(), 2);
                txtSalarioConyuge.Text = Aplicar_FormatoMoneda(Convert.ToDouble(dt.Rows[0]["SALARIO_CONYUGE"]), 1);
                txtDireccionTrabajoConyuge.Text = dt.Rows[0]["DIRECCION_TRABAJO_CONYUGE"].ToString();
                txtColoniaTrabajoConyuge.Text = dt.Rows[0]["COLONIA_CONYUGE"].ToString();
                cmbDepartamentoTrabajoConyuge.SelectedItem = dt.Rows[0]["DEPARTAMENTO_CONYUGE"].ToString();
                txtMunicipioTrabajoConyuge.Text = dt.Rows[0]["MUNICIPIO_CONYUGE"].ToString();
                txtZonaTrabajoConyuge.Text = dt.Rows[0]["ZONA_CONYUGE"].ToString();
                txtJefeInmediatoConyuge.Text = dt.Rows[0]["JEFE_INMEDIATO_CONYUGE"].ToString();
                txtTelefonoTrabajoConyuge.Text = dt.Rows[0]["TELEFONO_TRABAJO_CONYUGE"].ToString();
                txtTelefonoOficinaConyuge.Text = dt.Rows[0]["TELEFONO_OFICINA_CONYUGE"].ToString();
                txtExtensionConyuge.Text = dt.Rows[0]["TELEFONO_EXTENSION_CONYUGE"].ToString();
                txtTelefonoMovilConyuge.Text = dt.Rows[0]["TELEFONO_MOVIL_CONYUGE"].ToString();
                txtTotalIngresosConyuge.Text = Aplicar_FormatoMoneda(Convert.ToDouble(dt.Rows[0]["TOTAL_INGRESOS_CONYUGE"]), 1);

                txtNombreReferenciaPersonal1.Text = dt.Rows[0]["NOMBRE_REFERENCIA_PERSONAL1"].ToString();
                cmbRelacion1.SelectedItem = dt.Rows[0]["PARENTESCO_REFERENCIA_PERSONAL1"].ToString();
                txtTelefonoCasaReferenciaPersonal1.Text = dt.Rows[0]["TELEFONO_CASA_REFERENCIA_PERSONAL1"].ToString();
                txtTelefonoMovilReferenciaPersonal1.Text = dt.Rows[0]["TELEFONO_MOVIL_REFERENCIA_PERSONAL1"].ToString();
                txtTelefonoTrabajoReferenciaPersonal1.Text = dt.Rows[0]["TELEFONO_TRABAJO_REFERENCIA_PERSONAL1"].ToString();

                txtNombreReferenciaPersonal2.Text = dt.Rows[0]["NOMBRE_REFERENCIA_PERSONAL2"].ToString();
                cmbRelacion2.SelectedItem = dt.Rows[0]["PARENTESCO_REFERENCIA_PERSONAL2"].ToString();
                txtTelefonoCasaReferenciaPersonal2.Text = dt.Rows[0]["TELEFONO_CASA_REFERENCIA_PERSONAL2"].ToString();
                txtTelefonoMovilReferenciaPersonal2.Text = dt.Rows[0]["TELEFONO_MOVIL_REFERENCIA_PERSONAL2"].ToString();
                txtTelefonoTrabajoReferenciaPersonal2.Text = dt.Rows[0]["TELEFONO_TRABAJO_REFERENCIA_PERSONAL2"].ToString();

                txtNombreReferenciaPersonal3.Text = dt.Rows[0]["NOMBRE_REFERENCIA_PERSONAL3"].ToString();
                cmbRelacion3.SelectedItem = dt.Rows[0]["PARENTESCO_REFERENCIA_PERSONAL3"].ToString();
                txtTelefonoCasaReferenciaPersonal3.Text = dt.Rows[0]["TELEFONO_CASA_REFERENCIA_PERSONAL3"].ToString();
                txtTelefonoMovilReferenciaPersonal3.Text = dt.Rows[0]["TELEFONO_MOVIL_REFERENCIA_PERSONAL3"].ToString();
                txtTelefonoTrabajoReferenciaPersonal3.Text = dt.Rows[0]["TELEFONO_TRABAJO_REFERENCIA_PERSONAL3"].ToString();

                txtNombreReferenciaFamiliar1.Text = dt.Rows[0]["NOMBRE_REFERENCIA_FAMILIAR1"].ToString();
                cmbParentesco1.SelectedItem = dt.Rows[0]["PARENTESCO_REFERENCIA_FAMILIAR1"].ToString();
                txtTelefonoCasaReferenciaFamiliar1.Text = dt.Rows[0]["TELEFONO_CASA_REFERENCIA_FAMILIAR1"].ToString();
                txtTelefonoMovilReferenciaFamiliar1.Text = dt.Rows[0]["TELEFONO_MOVIL_REFERENCIA_FAMILIAR1"].ToString();
                txtTelefonoTrabajoReferenciaFamiliar1.Text = dt.Rows[0]["TELEFONO_TRABAJO_REFERENCIA_FAMILIAR1"].ToString();

                txtNombreReferenciaFamiliar2.Text = dt.Rows[0]["NOMBRE_REFERENCIA_FAMILIAR2"].ToString();
                cmbParentesco2.SelectedItem = dt.Rows[0]["PARENTESCO_REFERENCIA_FAMILIAR2"].ToString();
                txtTelefonoCasaReferenciaFamiliar2.Text = dt.Rows[0]["TELEFONO_CASA_REFERENCIA_FAMILIAR2"].ToString();
                txtTelefonoMovilReferenciaFamiliar2.Text = dt.Rows[0]["TELEFONO_MOVIL_REFERENCIA_FAMILIAR2"].ToString();
                txtTelefonoTrabajoReferenciaFamiliar2.Text = dt.Rows[0]["TELEFONO_TRABAJO_REFERENCIA_FAMILIAR2"].ToString();

                txtNombreReferenciaFamiliar3.Text = dt.Rows[0]["NOMBRE_REFERENCIA_FAMILIAR3"].ToString();
                cmbParentesco3.SelectedItem = dt.Rows[0]["PARENTESCO_REFERENCIA_FAMILIAR3"].ToString();
                txtTelefonoCasaReferenciaFamiliar3.Text = dt.Rows[0]["TELEFONO_CASA_REFERENCIA_FAMILIAR3"].ToString();
                txtTelefonoMovilReferenciaFamiliar3.Text = dt.Rows[0]["TELEFONO_MOVIL_REFERENCIA_FAMILIAR3"].ToString();
                txtTelefonoTrabajoReferenciaFamiliar3.Text = dt.Rows[0]["TELEFONO_TRABAJO_REFERENCIA_FAMILIAR3"].ToString();

                txtArticulo.Text = dt.Rows[0]["ARTICULO"].ToString();
                txtPrecio.Text = Aplicar_FormatoMoneda(Convert.ToDouble(dt.Rows[0]["PRECIO"]), 1);
                txtEnganche.Text = Aplicar_FormatoMoneda(Convert.ToDouble(dt.Rows[0]["ENGANCHE"]), 1);
                txtSaldo.Text = Aplicar_FormatoMoneda(Convert.ToDouble(dt.Rows[0]["SALDO"]), 1);
                txtPlazo.Text = dt.Rows[0]["PLAZO"].ToString();
                txtValorCuota.Text = Aplicar_FormatoMoneda(Convert.ToDouble(dt.Rows[0]["VALOR_CUOTA"]), 1);
                cmbTipoCredito.SelectedItem = dt.Rows[0]["TIPO_CREDITO"].ToString();
            }
        }

        // Obtiene todos los comentarios de la solicitud
        private void Obtener_Comentarios()
        {
            // Create the Grid    
            Grid DynamicGrid = new Grid();
            DynamicGrid.Width = 750;
            DynamicGrid.VerticalAlignment = VerticalAlignment.Top;
            DynamicGrid.Background = new SolidColorBrush(Colors.WhiteSmoke);

            // Create Columns    
            ColumnDefinition gridCol1 = new ColumnDefinition();
            gridCol1.Width = GridLength.Auto;

            ColumnDefinition gridCol2 = new ColumnDefinition();
            gridCol2.Width = GridLength.Auto;

            ColumnDefinition gridCol3 = new ColumnDefinition();
            gridCol3.Width = GridLength.Auto;

            ColumnDefinition gridCol4 = new ColumnDefinition();
            gridCol4.Width = GridLength.Auto;

            DynamicGrid.ColumnDefinitions.Add(gridCol1);
            DynamicGrid.ColumnDefinitions.Add(gridCol2);
            DynamicGrid.ColumnDefinitions.Add(gridCol3);
            DynamicGrid.ColumnDefinitions.Add(gridCol4);

            // Create Rows    
            RowDefinition gridRow1 = new RowDefinition();
            gridRow1.Height = new GridLength(25);
            RowDefinition gridRow2 = new RowDefinition();
            gridRow2.Height = new GridLength(25);
            RowDefinition gridRow3 = new RowDefinition();
            gridRow3.Height = new GridLength(25);
            RowDefinition gridRow4 = new RowDefinition();
            gridRow4.Height = new GridLength(25);
            RowDefinition gridRow5 = new RowDefinition();
            gridRow5.Height = new GridLength(25);
            RowDefinition gridRow6 = new RowDefinition();
            gridRow6.Height = new GridLength(25);
            RowDefinition gridRow7 = new RowDefinition();
            gridRow7.Height = new GridLength(25);
            RowDefinition gridRow8 = new RowDefinition();
            gridRow8.Height = new GridLength(25);
            RowDefinition gridRow9 = new RowDefinition();
            gridRow9.Height = new GridLength(25);
            RowDefinition gridRow10 = new RowDefinition();
            gridRow10.Height = new GridLength(25);
            RowDefinition gridRow11 = new RowDefinition();
            gridRow11.Height = new GridLength(25);
            RowDefinition gridRow12 = new RowDefinition();
            gridRow12.Height = new GridLength(25);
            RowDefinition gridRow13 = new RowDefinition();
            gridRow13.Height = new GridLength(25);
            RowDefinition gridRow14 = new RowDefinition();
            gridRow14.Height = new GridLength(25);
            RowDefinition gridRow15 = new RowDefinition();
            gridRow15.Height = new GridLength(25);

            DynamicGrid.RowDefinitions.Add(gridRow1);
            DynamicGrid.RowDefinitions.Add(gridRow2);
            DynamicGrid.RowDefinitions.Add(gridRow3);
            DynamicGrid.RowDefinitions.Add(gridRow4);
            DynamicGrid.RowDefinitions.Add(gridRow5);
            DynamicGrid.RowDefinitions.Add(gridRow6);
            DynamicGrid.RowDefinitions.Add(gridRow7);
            DynamicGrid.RowDefinitions.Add(gridRow8);
            DynamicGrid.RowDefinitions.Add(gridRow9);
            DynamicGrid.RowDefinitions.Add(gridRow10);
            DynamicGrid.RowDefinitions.Add(gridRow11);
            DynamicGrid.RowDefinitions.Add(gridRow12);
            DynamicGrid.RowDefinitions.Add(gridRow13);
            DynamicGrid.RowDefinitions.Add(gridRow14);
            DynamicGrid.RowDefinitions.Add(gridRow15);

            // Add first column header    
            TextBlock txtBlock1 = new TextBlock();
            txtBlock1.Text = "ESTADO ";
            txtBlock1.FontSize = 13;
            txtBlock1.FontWeight = FontWeights.Bold;
            txtBlock1.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock1.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock1.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock1, 0);
            Grid.SetColumn(txtBlock1, 0);

            TextBlock txtBlock2 = new TextBlock();
            txtBlock2.Text = "FECHA ";
            txtBlock2.FontSize = 13;
            txtBlock2.FontWeight = FontWeights.Bold;
            txtBlock2.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock2.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock2.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock2, 0);
            Grid.SetColumn(txtBlock2, 1);

            TextBlock txtBlock3 = new TextBlock();
            txtBlock3.Text = "USUARIO";
            txtBlock3.FontSize = 13;
            txtBlock3.FontWeight = FontWeights.Bold;
            txtBlock3.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock3.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock3.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock3, 0);
            Grid.SetColumn(txtBlock3, 2);

            TextBlock txtBlock4 = new TextBlock();
            txtBlock4.Text = "COMENTARIOS";
            txtBlock4.FontSize = 13;
            txtBlock4.FontWeight = FontWeights.Bold;
            txtBlock4.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock4.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock4.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock4, 0);
            Grid.SetColumn(txtBlock4, 3);

            //// Add column headers to the Grid    
            DynamicGrid.Children.Add(txtBlock1);
            DynamicGrid.Children.Add(txtBlock2);
            DynamicGrid.Children.Add(txtBlock3);
            DynamicGrid.Children.Add(txtBlock4);

            if (txtSolicitud.Text=="") { txtSolicitud.Text = "0"; }

            string strSql = @"select id_solicitud, case estado when 1 then 'Ingresado' when 2 then 'Enviado a Créditos' when 99 then 'Aprobado' when 98 then 'Retornado' when 97 then 'Rechazado'end as estado, FECHA, USUARIO, COMENTARIOS 
                            from SI_Estados where ID_SOLICITUD = "+ txtSolicitud.Text+" order by id desc";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUTILS))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(strSql, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            // Dibujar el detalle
            if (dt.Rows.Count > 0)
            {
                for (int i = 1; i <= dt.Rows.Count; i++)
                {
                    // Create first Row    
                    TextBlock txtbSolicitud = new TextBlock();
                    txtbSolicitud.Text = dt.Rows[i - 1]["estado"].ToString();
                    txtbSolicitud.FontSize = 12;
                    txtbSolicitud.Width = 120;
                    txtbSolicitud.HorizontalAlignment = HorizontalAlignment.Left;
                    Grid.SetRow(txtbSolicitud, i);
                    Grid.SetColumn(txtbSolicitud, 0);

                    TextBlock txtbCliente = new TextBlock();
                    txtbCliente.Text = Convertir_Fecha(dt.Rows[i - 1]["fecha"].ToString(),2);
                    txtbCliente.FontSize = 12;
                    txtbCliente.Width = 120;
                    txtbCliente.HorizontalAlignment = HorizontalAlignment.Right;
                    Grid.SetRow(txtbCliente, i);
                    Grid.SetColumn(txtbCliente, 1);

                    TextBlock txtbDpi = new TextBlock();
                    txtbDpi.Text = dt.Rows[i - 1]["usuario"].ToString();
                    txtbDpi.FontSize = 12;
                    txtbDpi.Width = 100;
                    txtbDpi.HorizontalAlignment = HorizontalAlignment.Right;
                    Grid.SetRow(txtbDpi, i);
                    Grid.SetColumn(txtbDpi, 2);

                    TextBlock txtbArticulo = new TextBlock();
                    txtbArticulo.Text = dt.Rows[i - 1]["comentarios"].ToString();
                    txtbArticulo.FontSize = 12;
                    txtbArticulo.Width = 390;
                    txtbArticulo.HorizontalAlignment = HorizontalAlignment.Left;
                    Grid.SetRow(txtbArticulo, i);
                    Grid.SetColumn(txtbArticulo, 3);

                    DynamicGrid.Children.Add(txtbSolicitud);
                    DynamicGrid.Children.Add(txtbCliente);
                    DynamicGrid.Children.Add(txtbDpi);
                    DynamicGrid.Children.Add(txtbArticulo);
                }
            }

            // Display grid into a Window    
            ScrollComentarios.Content = DynamicGrid;
        }

        // Obtiene el detalle de las solicitudes
        private void Obtener_Solicitudes()
        {
            // Create the Grid    
            Grid DynamicGrid = new Grid();
            DynamicGrid.Width = 750;
            DynamicGrid.VerticalAlignment = VerticalAlignment.Top;
            DynamicGrid.Background = new SolidColorBrush(Colors.WhiteSmoke);

            // Create Columns    
            ColumnDefinition gridCol1 = new ColumnDefinition();
            gridCol1.Width = GridLength.Auto;

            ColumnDefinition gridCol2 = new ColumnDefinition();
            gridCol2.Width = GridLength.Auto;

            ColumnDefinition gridCol3 = new ColumnDefinition();
            gridCol3.Width = GridLength.Auto;

            ColumnDefinition gridCol4 = new ColumnDefinition();
            gridCol4.Width = GridLength.Auto;

            ColumnDefinition gridCol5 = new ColumnDefinition();
            gridCol5.Width = GridLength.Auto;

            ColumnDefinition gridCol6 = new ColumnDefinition();
            gridCol6.Width = GridLength.Auto;

            ColumnDefinition gridCol7 = new ColumnDefinition();
            gridCol7.Width = GridLength.Auto;

            DynamicGrid.ColumnDefinitions.Add(gridCol1);
            DynamicGrid.ColumnDefinitions.Add(gridCol2);
            DynamicGrid.ColumnDefinitions.Add(gridCol3);
            DynamicGrid.ColumnDefinitions.Add(gridCol4);
            DynamicGrid.ColumnDefinitions.Add(gridCol5);
            DynamicGrid.ColumnDefinitions.Add(gridCol6);
            DynamicGrid.ColumnDefinitions.Add(gridCol7);

            // Create Rows    
            RowDefinition gridRow1 = new RowDefinition();
            gridRow1.Height = new GridLength(25);
            RowDefinition gridRow2 = new RowDefinition();
            gridRow2.Height = new GridLength(25);
            RowDefinition gridRow3 = new RowDefinition();
            gridRow3.Height = new GridLength(25);
            RowDefinition gridRow4 = new RowDefinition();
            gridRow4.Height = new GridLength(25);
            RowDefinition gridRow5 = new RowDefinition();
            gridRow5.Height = new GridLength(25);
            RowDefinition gridRow6 = new RowDefinition();
            gridRow6.Height = new GridLength(25);
            RowDefinition gridRow7 = new RowDefinition();
            gridRow7.Height = new GridLength(25);
            RowDefinition gridRow8 = new RowDefinition();
            gridRow8.Height = new GridLength(25);
            RowDefinition gridRow9 = new RowDefinition();
            gridRow9.Height = new GridLength(25);
            RowDefinition gridRow10 = new RowDefinition();
            gridRow10.Height = new GridLength(25);
            RowDefinition gridRow11 = new RowDefinition();
            gridRow11.Height = new GridLength(25);
            RowDefinition gridRow12 = new RowDefinition();
            gridRow12.Height = new GridLength(25);
            RowDefinition gridRow13 = new RowDefinition();
            gridRow13.Height = new GridLength(25);
            RowDefinition gridRow14 = new RowDefinition();
            gridRow14.Height = new GridLength(25);
            RowDefinition gridRow15 = new RowDefinition();
            gridRow15.Height = new GridLength(25);

            DynamicGrid.RowDefinitions.Add(gridRow1);
            DynamicGrid.RowDefinitions.Add(gridRow2);
            DynamicGrid.RowDefinitions.Add(gridRow3);
            DynamicGrid.RowDefinitions.Add(gridRow4);
            DynamicGrid.RowDefinitions.Add(gridRow5);
            DynamicGrid.RowDefinitions.Add(gridRow6);
            DynamicGrid.RowDefinitions.Add(gridRow7);
            DynamicGrid.RowDefinitions.Add(gridRow8);
            DynamicGrid.RowDefinitions.Add(gridRow9);
            DynamicGrid.RowDefinitions.Add(gridRow10);
            DynamicGrid.RowDefinitions.Add(gridRow11);
            DynamicGrid.RowDefinitions.Add(gridRow12);
            DynamicGrid.RowDefinitions.Add(gridRow13);
            DynamicGrid.RowDefinitions.Add(gridRow14);
            DynamicGrid.RowDefinitions.Add(gridRow15);

            // Add first column header    
            TextBlock txtBlock1 = new TextBlock();
            txtBlock1.Text = "SOLICITUD ";
            txtBlock1.FontSize = 13;
            txtBlock1.FontWeight = FontWeights.Bold;
            txtBlock1.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock1.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock1.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock1, 0);
            Grid.SetColumn(txtBlock1, 0);

            TextBlock txtBlock2 = new TextBlock();
            txtBlock2.Text = "CLIENTE ";
            txtBlock2.FontSize = 13;
            txtBlock2.FontWeight = FontWeights.Bold;
            txtBlock2.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock2.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock2.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock2, 0);
            Grid.SetColumn(txtBlock2, 1);

            TextBlock txtBlock3 = new TextBlock();
            txtBlock3.Text = "DPI";
            txtBlock3.FontSize = 13;
            txtBlock3.FontWeight = FontWeights.Bold;
            txtBlock3.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock3.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock3.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock3, 0);
            Grid.SetColumn(txtBlock3, 2);

            TextBlock txtBlock4 = new TextBlock();
            txtBlock4.Text = "ARTÍCULO";
            txtBlock4.FontSize = 13;
            txtBlock4.FontWeight = FontWeights.Bold;
            txtBlock4.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock4.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock4.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock4, 0);
            Grid.SetColumn(txtBlock4, 3);

            TextBlock txtBlock5 = new TextBlock();
            txtBlock5.Text = "VALOR";
            txtBlock5.FontSize = 13;
            txtBlock5.FontWeight = FontWeights.Bold;
            txtBlock5.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock5.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock5.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock5, 0);
            Grid.SetColumn(txtBlock5, 4);

            TextBlock txtBlock6 = new TextBlock();
            txtBlock6.Text = "ESTADO";
            txtBlock6.FontSize = 13;
            txtBlock6.FontWeight = FontWeights.Bold;
            txtBlock6.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock6.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock6.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock6, 0);
            Grid.SetColumn(txtBlock6, 5);

            TextBlock txtBlock7 = new TextBlock();
            txtBlock7.Text = "ACCIÓN";
            txtBlock7.FontSize = 13;
            txtBlock7.FontWeight = FontWeights.Bold;
            txtBlock7.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            txtBlock7.Background = new SolidColorBrush(Colors.DarkBlue);
            txtBlock7.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(txtBlock7, 0);
            Grid.SetColumn(txtBlock7, 6);

            //// Add column headers to the Grid    
            DynamicGrid.Children.Add(txtBlock1);
            DynamicGrid.Children.Add(txtBlock2);
            DynamicGrid.Children.Add(txtBlock3);
            DynamicGrid.Children.Add(txtBlock4);
            DynamicGrid.Children.Add(txtBlock5);
            DynamicGrid.Children.Add(txtBlock6);
            DynamicGrid.Children.Add(txtBlock7);

            string strSql = "";
            
            switch (Rol)
            {
                case "V":
                    strSql = @"select SO.ID, SO.PRIMER_APELLIDO + ' ' + SO.SEGUNDO_APELLIDO + ' '+ SO.PRIMER_NOMBRE AS CLIENTE, sO.DPI_PASAPORTE_CLIENTE, 
                            SO.ARTICULO, SO.PRECIO, 
                            case (select top 1 ESTADO from SI_Estados where ID_SOLICITUD = SO.ID order by id desc) 
                            when 1 then 'Ingresado'  
                            when 2 then 'Enviado a Créditos' 
                            when 99 then 'Aprobado'
                            when 98 then 'Retornado'
                            when 97 then 'Rechazado'
                            end as ESTADO 
                            FROM SI_Solicitudes SO
                            where SO.USUARIO = '" + txtUsuario.Text + "'";

                    break;

                case "A":
                    strSql = @"select * from (
                            select SO.ID, SO.PRIMER_APELLIDO + ' ' + SO.SEGUNDO_APELLIDO + ' '+ SO.PRIMER_NOMBRE AS CLIENTE, sO.DPI_PASAPORTE_CLIENTE, 
                            SO.ARTICULO, SO.PRECIO, 
                            case (select top 1 ESTADO from SI_Estados where ID_SOLICITUD = SO.ID order by id desc) 
                            when 1 then 'Ingresado'  
                            when 2 then 'Enviado a Créditos' 
                            when 99 then 'Aprobado'
                            when 98 then 'Retornado'
                            when 97 then 'Rechazado'
                            end as ESTADO 
                            FROM SI_Solicitudes SO) as datos
                            where estado = 'Enviado a Créditos'";
                    
                    break;
            }

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUTILS))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(strSql, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            // Dibujar el detalle
            if (dt.Rows.Count > 0)
            {
                for (int i = 1; i <= dt.Rows.Count; i++)
                {
                    // Create first Row    
                    TextBlock txtbSolicitud = new TextBlock();
                    txtbSolicitud.Text = dt.Rows[i - 1]["id"].ToString();
                    txtbSolicitud.FontSize = 12;
                    txtbSolicitud.Width = 60;
                    txtbSolicitud.HorizontalAlignment = HorizontalAlignment.Center;
                    Grid.SetRow(txtbSolicitud, i);
                    Grid.SetColumn(txtbSolicitud, 0);

                    TextBlock txtbCliente = new TextBlock();
                    txtbCliente.Text = dt.Rows[i - 1]["cliente"].ToString();
                    txtbCliente.FontSize = 12;
                    txtbCliente.Width = 175;
                    txtbCliente.HorizontalAlignment = HorizontalAlignment.Left;
                    Grid.SetRow(txtbCliente, i);
                    Grid.SetColumn(txtbCliente, 1);

                    TextBlock txtbDpi = new TextBlock();
                    txtbDpi.Text = dt.Rows[i - 1]["dpi_pasaporte_cliente"].ToString();
                    txtbDpi.FontSize = 12;
                    txtbDpi.Width = 100;
                    txtbDpi.HorizontalAlignment = HorizontalAlignment.Left;
                    Grid.SetRow(txtbDpi, i);
                    Grid.SetColumn(txtbDpi, 2);

                    TextBlock txtbArticulo = new TextBlock();
                    txtbArticulo.Text = dt.Rows[i - 1]["articulo"].ToString();
                    txtbArticulo.FontSize = 12;
                    txtbArticulo.Width = 130;
                    txtbArticulo.HorizontalAlignment = HorizontalAlignment.Left;
                    Grid.SetRow(txtbArticulo, i);
                    Grid.SetColumn(txtbArticulo, 3);

                    TextBlock txtbValor = new TextBlock();
                    txtbValor.Text = Aplicar_FormatoMoneda(Convert.ToDouble(dt.Rows[i - 1]["precio"]), 1);
                    txtbValor.FontSize = 12;
                    txtbValor.HorizontalAlignment = HorizontalAlignment.Right;
                    txtbValor.Width = 80;
                    Grid.SetRow(txtbValor, i);
                    Grid.SetColumn(txtbValor, 4);

                    TextBlock txtbEstado = new TextBlock();
                    txtbEstado.Text = dt.Rows[i - 1]["estado"].ToString();
                    txtbEstado.FontSize = 12;
                    txtbEstado.FontWeight = FontWeights.Bold;
                    txtbEstado.Width = 120;
                    txtbEstado.HorizontalAlignment = HorizontalAlignment.Left;
                    Grid.SetRow(txtbEstado, i);
                    Grid.SetColumn(txtbEstado, 5);

                    Button btnDetalleAccion = new Button();
                    btnDetalleAccion.Content = "VER";
                    btnDetalleAccion.Width = 60;
                    btnDetalleAccion.Click += btnDetalleAccion_Click;
                    btnDetalleAccion.Tag = dt.Rows[i - 1]["id"].ToString();
                    Grid.SetRow(btnDetalleAccion, i);
                    Grid.SetColumn(btnDetalleAccion, 6);

                    DynamicGrid.Children.Add(btnDetalleAccion);

                    DynamicGrid.Children.Add(txtbSolicitud);
                    DynamicGrid.Children.Add(txtbCliente);
                    DynamicGrid.Children.Add(txtbDpi);
                    DynamicGrid.Children.Add(txtbArticulo);
                    DynamicGrid.Children.Add(txtbValor);
                    DynamicGrid.Children.Add(txtbEstado);
                }
            }

            // Display grid into a Window    
            ScrollSolicitudes.Content = DynamicGrid;
        }

        // Valida correo electronico
        private bool Validar_CorreoElectronico(string strCorreoElectronico)
        {
            string strExpresion;
            strExpresion = "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";
            if (Regex.IsMatch(strCorreoElectronico, strExpresion))
            {
                if (Regex.Replace(strCorreoElectronico, strExpresion, String.Empty).Length == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        // Valida fecha
        private bool Validar_Fecha(string strFecha)
        {
            string strExpresion;
            strExpresion = "^([0]?[0-9]|[12][0-9]|[3][01])[./-]([0]?[1-9]|[1][0-2])[./-]([0-9]{4}|[0-9]{2})$";
            if (Regex.IsMatch(strFecha, strExpresion))
            {
                if (Regex.Replace(strFecha, strExpresion, String.Empty).Length == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        // Convierte la fecha para ser almacenada en DB
        private string Convertir_Fecha(string strFecha, int intTipo)
        {
            switch (intTipo)
            {
                case 1: // 2021-01-01
                    if (strFecha == "") { return "NULL"; }
                    else { return "'" + strFecha.Substring(6, 4) + "-" + strFecha.Substring(3, 2) + "-" + strFecha.Substring(0, 2) + "'"; }
                case 2: // 01/01/2021
                    if (strFecha != "NULL" & strFecha !="") { return strFecha.Substring(0,10); }
                    else { return strFecha;  }
                default:
                    return "";
            }
        }

        // Convierte la moneda para ser almacenada en DB
        private double Convertir_Moneda(string strMoneda)
        {
            if (strMoneda == "") { return 0.00;}
            else { return Convert.ToDouble(strMoneda); }
        }

        // Convierte el comobo para ser almacenado en DB
        private string Convertir_Combo(string strCombo)
        {
            if (strCombo.IndexOf("System.Windows.Controls.ComboBoxItem: ") != -1) { return strCombo.Replace("System.Windows.Controls.ComboBoxItem: ", ""); }
            else { return strCombo; }
        }

        // Valida que solo se ingrese números
        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex objRegex = new Regex("[^0-9]+");
            e.Handled = objRegex.IsMatch(e.Text);
        }

        // Evento del boton ABC
        private void btnABC_Click(object sender, RoutedEventArgs e)
        {
            switch (btnABC.Content.ToString())
            {
                case "Nuevo":
                    Habilitar_Controles();

                    MessageBox.Show("Ahora puede ingresar una nueva solicitud!", Sistema, MessageBoxButton.OK, MessageBoxImage.Information);

                    btnABC.Content = "Guardar";

                    tabSolicitud.Focus();

                    break;

                case "Guardar":
                    var objResultado1 = MessageBox.Show("Desea guardar la solicitud?", Sistema, MessageBoxButton.YesNo, MessageBoxImage.Question);
                    string strError1 = "";

                    if (objResultado1 == MessageBoxResult.Yes)
                    {
                        if (Guardar_Solicitud(ref strError1))
                        {
                            Guardar_Estado(Obtener_UltimoSolicitud(), 1, ref strError1);

                            MessageBox.Show("La solicitud ha sido guardada satisfactoriamente!", Sistema, MessageBoxButton.OK, MessageBoxImage.Information);

                            Limpiar_Controles();

                            Bloquear_Controles();

                            Obtener_Solicitudes();

                            tabSolicitudes.Focus();

                            btnABC.Content = "Nuevo";
                        }
                        else { MessageBox.Show("La solicitud no fue guardada! " + "Error: " + strError1, Sistema, MessageBoxButton.OK, MessageBoxImage.Error); }
                    }

                    break;

                case "Habilitar":
                    Habilitar_Controles();

                    btnABC.Content = "Modificar";

                    MessageBox.Show("Ahora puede modificar la solicitud!", Sistema, MessageBoxButton.OK, MessageBoxImage.Information);

                    break;

                case "Modificar":
                    var objResultado2 = MessageBox.Show("Desea modificar la solicitud?", Sistema, MessageBoxButton.YesNo, MessageBoxImage.Question);
                    string strError2 = "";

                    if (objResultado2 == MessageBoxResult.Yes)
                    {
                        if (Modificar_Solicitud(ref strError2))
                        {
                            Guardar_Estado(Obtener_UltimoSolicitud(), 1, ref strError2);

                            MessageBox.Show("La solicitud ha sido modificada satisfactoriamente!", Sistema, MessageBoxButton.OK, MessageBoxImage.Information);

                            Limpiar_Controles();

                            Bloquear_Controles();

                            Obtener_Solicitudes();

                            tabSolicitudes.Focus();

                            btnABC.Content = "Nuevo";
                        }
                        else { MessageBox.Show("La solicitud no fue modificada! " + "Error: " + strError2, Sistema, MessageBoxButton.OK, MessageBoxImage.Error); }
                    }

                    btnABC.Content = "Nuevo";
                    break;

                default:
                    break;
            }
        }

        // Evento para aplicar permisos por Rol
        private void Aplicar_PermisosRol()
        {
            txtFechaSolicitud.IsEnabled = false;
            txtFechaSolicitud.Text = DateTime.Today.ToString();
            //cmbVendedor.Items.Clear();

            switch (Rol)
            {
                case "V":
                    btnABC.IsEnabled = true;
                    btnABC.Content = "Nuevo";
                    cmbVendedor.Items.Insert(0, Vendedor);
                    cmbVendedor.SelectedIndex = 0;

                    break;

                case "A":
                    lblVendedor.Content = "Analista:";
                    lblSolicitud.Visibility = Visibility.Hidden;
                    txtSolicitud.Visibility = Visibility.Hidden;
                    lblAccion.Visibility = Visibility.Visible;
                    cmbAccion.Visibility = Visibility.Visible;
                    btnAccion.Content = "ACTUALIZAR";
                    btnAccion.IsEnabled = true;
                    txtComentarios.IsEnabled = true;

                    break;

                default:

                    break;
            }
        }

        // Evento del boton login
        private void btnLogin_Click(object sender, RoutedEventArgs e)
        {
            // Si el boton esta en modo Login
            if (btnLogin.Content.ToString() == "Login")
            {
                if (Validar_Credenciales())
                {
                    // Deshabilita los controles y login
                    txtUsuario.IsEnabled = false;
                    txtClave.IsEnabled = false;
                    btnLogin.Content = "Logout";

                    // Dirige el foco al formulario de datos del cliente
                    Habilitar_Tabuladores();
                    Llenar_Combos();
                    Obtener_Solicitudes();
                    Bloquear_Controles();
                    tabSolicitudes.Focus();
                    btnCancelar.IsEnabled = true;
                    Aplicar_PermisosRol();
                }
                else { MessageBox.Show("Debe de ingresar sus credenciales de STOD para continuar!", Sistema, MessageBoxButton.OK, MessageBoxImage.Error); }
            }
            else
            {
                // Habilita los controles y logout
                txtUsuario.IsEnabled = true;
                txtUsuario.Text = "";
                txtClave.IsEnabled = true;
                txtClave.Password = "";
                btnLogin.Content = "Login";
                lblnombreusuario.Visibility = Visibility.Hidden;
                lblareausuario.Visibility = Visibility.Hidden;

                // Bloquea la aplicación
                Limpiar_Controles();

                Bloquear_Controles();

                Bloquear_Tabuladores();

                tabAcceso.Focus();
            }
        }

        // Metodo que obtiene datos de los estados de la solicitud
        private void Obtener_DatosEstadoSolicitud(ref int refEstado, ref string refFecha, ref string refUsuario, ref string refComentarios)
        {
            string strSql = "select top 1 ESTADO, fecha, usuario, comentarios from SI_Estados where ID_SOLICITUD = 4 order by id desc";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUTILS))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(strSql, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            if (dt.Rows.Count > 0) 
            { 
                refEstado =  Convert.ToInt32(dt.Rows[0]["estado"]);
                refFecha = dt.Rows[0]["fecha"].ToString();
                refUsuario = dt.Rows[0]["usuario"].ToString();
                refComentarios = dt.Rows[0]["comentarios"].ToString();
            }
        }

        // Metodo que obtiene la ultima solicitud por usuario
        private int Obtener_UltimoSolicitud()
        {
            string strSql = "select top 1 id from si_solicitudes where usuario = '"+txtUsuario.Text+"' order by id desc";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUTILS))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(strSql, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            if (dt.Rows.Count > 0) { return Convert.ToInt32(dt.Rows[0]["id"]); }
            else { return 0; }
        }

        // Metodo para validar el rol
        private bool Validar_Rol(string strUsuario)
        {
            string strSql = @"select * from SI_Usuarios where USUARIO_STOD = '" + strUsuario + "'";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUTILS))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(strSql, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            if (dt.Rows.Count > 0)
            {
                Rol = dt.Rows[0]["rol"].ToString();
                Vendedor = dt.Rows[0]["nombre"].ToString();
                return true;
            }
            else { return false; }
        }

        // Metodo para validar las credenciales
        private bool Validar_Credenciales()
        {
            string strSql = @"select U.UserId, U.UserName, M.Password, M.PasswordSalt, M.Email from aspnet_Users U, aspnet_Membership M where M.UserId = U.UserId and upper(U.UserName) = '" + txtUsuario.Text.ToUpper() + "'";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringSTOD))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(strSql, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["password"].ToString() == EncodePassword(txtClave.Password, dt.Rows[0]["passwordsalt"].ToString()))
                {
                    // Revisa si existe en la tabla de usuarios de Score Interno
                    if (Validar_Rol(txtUsuario.Text)) 
                    {
                        // Coloca el nombre del usuario y el correo
                        lblnombreusuario.Content = dt.Rows[0]["username"].ToString();
                        lblnombreusuario.Visibility = Visibility.Visible;

                        lblareausuario.Visibility = Visibility.Visible;
                        switch (Rol)
                        {
                            case "V":
                                lblareausuario.Content = "Vendedor";
                                break;
                            case "A":
                                lblareausuario.Content = "Analista";
                                break;
                            default:
                                break;
                        }

                        return true; 
                    }
                    else { return false; }
                }
                else { return false; }
            }
            else { return false; }
        }

        // Metodo para validar el hash del password del usuario de STOD
        public string EncodePassword(string pass, string saltBase64)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(pass);
            byte[] src = Convert.FromBase64String(saltBase64);
            byte[] dst = new byte[src.Length + bytes.Length];
            Buffer.BlockCopy(src, 0, dst, 0, src.Length);
            Buffer.BlockCopy(bytes, 0, dst, src.Length, bytes.Length);
            HashAlgorithm algorithm = HashAlgorithm.Create("SHA1");
            byte[] inArray = algorithm.ComputeHash(dst);
            return Convert.ToBase64String(inArray);
        }

        // Metodo que inicia la pantalla
        private void Iniciar_Aplicacion()
        {
            // Inicia la pantalla y coloca el foco en el formulario de seguridad
            Bloquear_Tabuladores();
            tabAcceso.Focus();
        }

        // Metodo para limpiar los controles
        private void Limpiar_Controles()
        {
            cmbMarca.SelectedIndex = 0;
            cmbTienda.SelectedIndex = 0;
            txtTelefonoVendedor.Text = "";
            cmbTipoCliente.SelectedIndex = 0;
            txtPrimerNombreCliente.Text = "";
            txtSegundoNombreCliente.Text = "";
            txtPrimerApellidoCliente.Text = "";
            txtSegundoApellidoCliente.Text = "";
            txtApellidoCasadaCliente.Text = "";
            txtDpi.Text = "";
            txtNacionalidad.Text = "";
            txtNit.Text = "";
            txtFechaNacimiento.Text = "";
            txtEstadoCivil.Text = "";
            txtSexo.Text = "";
            txtDependientes.Text = "";
            txtProfesion.Text = "";
            txtDireccion.Text = "";
            txtColonia.Text = "";
            cmbDepartamento.SelectedIndex = 0;
            txtMunicipio.Text = "";
            txtZona.Text = "";
            cmbTipoVivienda.SelectedIndex = 0;
            txtTelefonoCasa.Text = "";
            txtTelefonoMovil.Text = "";
            txtCorreoElectronico.Text = "";
            txtNombreEmpresaCliente.Text = "";
            txtCargo.Text = "";
            txtFechaIngreso.Text = "";
            txtIngresosFijos.Text = "";
            txtIngresosVariables.Text = "";
            txtOtrosIngresos.Text = "";
            txtFuenteOtrosIngresos.Text = "";
            txtDireccionTrabajoCliente.Text = "";
            txtColoniaTrabajoCliente.Text = "";
            cmbDepartamentoTrabajoCliente.SelectedIndex = 0;
            txtMunicipioTrabajoCliente.Text = "";
            txtZonaTrabajoCliente.Text = "";
            txtJefeInmediato.Text = "";
            txtTelefonoTrabajo.Text = "";
            txtTelefonoJefe.Text = "";
            txtExtension.Text = "";
            txtNitEmpresa.Text = "";
            txtTotalIngresos.Text = "";
            txtNombresApellidosConyuge.Text = "";
            txtDpiConyuge.Text = "";
            txtFechaNacimientoConyuge.Text = "";
            txtTrabajoConyuge.Text = "";
            txtCargoConyuge.Text = "";
            txtFechaIngresoConyuge.Text = "";
            txtSalarioConyuge.Text = "";
            txtDireccionTrabajoConyuge.Text = "";
            txtColoniaTrabajoConyuge.Text = "";
            cmbDepartamentoTrabajoConyuge.SelectedIndex = 0;
            txtMunicipioTrabajoConyuge.Text = "";
            txtZonaTrabajoConyuge.Text = "";
            txtJefeInmediatoConyuge.Text = "";
            txtTelefonoTrabajoConyuge.Text = "";
            txtTelefonoOficinaConyuge.Text = "";
            txtExtensionConyuge.Text = "";
            txtTelefonoMovilConyuge.Text = "";
            txtTotalIngresosConyuge.Text = "";
            txtNombreReferenciaPersonal1.Text = "";
            cmbRelacion1.SelectedIndex = 0;
            txtTelefonoCasaReferenciaPersonal1.Text = "";
            txtTelefonoMovilReferenciaPersonal1.Text = "";
            txtTelefonoTrabajoReferenciaPersonal1.Text = "";
            txtNombreReferenciaPersonal2.Text = "";
            cmbRelacion2.SelectedIndex = 0;
            txtTelefonoCasaReferenciaPersonal2.Text = "";
            txtTelefonoMovilReferenciaPersonal2.Text = "";
            txtTelefonoTrabajoReferenciaPersonal2.Text = "";
            txtNombreReferenciaPersonal3.Text = "";
            cmbRelacion3.SelectedIndex = 0;
            txtTelefonoCasaReferenciaPersonal3.Text = "";
            txtTelefonoMovilReferenciaPersonal3.Text = "";
            txtTelefonoTrabajoReferenciaPersonal3.Text = "";
            txtNombreReferenciaFamiliar1.Text = "";
            cmbParentesco1.SelectedIndex = 0;
            txtTelefonoCasaReferenciaFamiliar1.Text = "";
            txtTelefonoMovilReferenciaFamiliar1.Text = "";
            txtTelefonoTrabajoReferenciaFamiliar1.Text = "";
            txtNombreReferenciaFamiliar2.Text = "";
            cmbParentesco2.SelectedIndex = 0;
            txtTelefonoCasaReferenciaFamiliar2.Text = "";
            txtTelefonoMovilReferenciaFamiliar2.Text = "";
            txtTelefonoTrabajoReferenciaFamiliar2.Text = "";
            txtNombreReferenciaFamiliar3.Text = "";
            cmbParentesco3.SelectedIndex = 0;
            txtTelefonoCasaReferenciaFamiliar3.Text = "";
            txtTelefonoMovilReferenciaFamiliar3.Text = "";
            txtTelefonoTrabajoReferenciaFamiliar3.Text = "";
            txtArticulo.Text = "";
            txtPrecio.Text = "";
            txtPlazo.Text = "";
            txtEnganche.Text = "";
            txtValorCuota.Text = "";
            txtSaldo.Text = "";
            cmbTipoCredito.SelectedIndex = 0;
            lblMensajes.Content = "";
            txtComentarios.Text = "";
            txtEstadoSolicitud.Text = "";
            txtSolicitud.Text = "";
            Obtener_Comentarios();
        }

        // Metodo para habilitar controles
        private void Habilitar_Controles()
        {
            cmbMarca.IsEnabled = true;
            cmbTienda.IsEnabled = true;
            txtTelefonoVendedor.IsEnabled = true;
            cmbTipoCliente.IsEnabled = true;
            txtPrimerNombreCliente.IsEnabled = true;
            txtSegundoNombreCliente.IsEnabled = true;
            txtPrimerApellidoCliente.IsEnabled = true;
            txtSegundoApellidoCliente.IsEnabled = true;
            txtApellidoCasadaCliente.IsEnabled = true;
            txtDpi.IsEnabled = true;
            txtNacionalidad.IsEnabled = true;
            txtNit.IsEnabled = true;
            txtFechaNacimiento.IsEnabled = true;
            txtEstadoCivil.IsEnabled = true;
            txtSexo.IsEnabled = true;
            txtDependientes.IsEnabled = true;
            txtProfesion.IsEnabled = true;
            txtDireccion.IsEnabled = true;
            txtColonia.IsEnabled = true;
            cmbDepartamento.IsEnabled = true;
            txtMunicipio.IsEnabled = true;
            txtZona.IsEnabled = true;
            cmbTipoVivienda.IsEnabled = true;
            txtTelefonoCasa.IsEnabled = true;
            txtTelefonoMovil.IsEnabled = true;
            txtCorreoElectronico.IsEnabled = true;
            txtNombreEmpresaCliente.IsEnabled = true;
            txtCargo.IsEnabled = true;
            txtFechaIngreso.IsEnabled = true;
            txtIngresosFijos.IsEnabled = true;
            txtIngresosVariables.IsEnabled = true;
            txtOtrosIngresos.IsEnabled = true;
            txtFuenteOtrosIngresos.IsEnabled = true;
            txtDireccionTrabajoCliente.IsEnabled = true;
            txtColoniaTrabajoCliente.IsEnabled = true;
            cmbDepartamentoTrabajoCliente.IsEnabled = true;
            txtMunicipioTrabajoCliente.IsEnabled = true;
            txtZonaTrabajoCliente.IsEnabled = true;
            txtJefeInmediato.IsEnabled = true;
            txtTelefonoTrabajo.IsEnabled = true;
            txtTelefonoJefe.IsEnabled = true;
            txtExtension.IsEnabled = true;
            txtNitEmpresa.IsEnabled = true;
            txtTotalIngresos.IsEnabled = true;
            txtNombresApellidosConyuge.IsEnabled = true;
            txtDpiConyuge.IsEnabled = true;
            txtFechaNacimientoConyuge.IsEnabled = true;
            txtTrabajoConyuge.IsEnabled = true;
            txtCargoConyuge.IsEnabled = true;
            txtFechaIngresoConyuge.IsEnabled = true;
            txtSalarioConyuge.IsEnabled = true;
            txtDireccionTrabajoConyuge.IsEnabled = true;
            txtColoniaTrabajoConyuge.IsEnabled = true;
            cmbDepartamentoTrabajoConyuge.IsEnabled = true;
            txtMunicipioTrabajoConyuge.IsEnabled = true;
            txtZonaTrabajoConyuge.IsEnabled = true;
            txtJefeInmediatoConyuge.IsEnabled = true;
            txtTelefonoTrabajoConyuge.IsEnabled = true;
            txtTelefonoOficinaConyuge.IsEnabled = true;
            txtExtensionConyuge.IsEnabled = true;
            txtTelefonoMovilConyuge.IsEnabled = true;
            txtTotalIngresosConyuge.IsEnabled = true;
            txtNombreReferenciaPersonal1.IsEnabled = true;
            cmbRelacion1.IsEnabled = true;
            txtTelefonoCasaReferenciaPersonal1.IsEnabled = true;
            txtTelefonoMovilReferenciaPersonal1.IsEnabled = true;
            txtTelefonoTrabajoReferenciaPersonal1.IsEnabled = true;
            txtNombreReferenciaPersonal2.IsEnabled = true;
            cmbRelacion2.IsEnabled = true;
            txtTelefonoCasaReferenciaPersonal2.IsEnabled = true;
            txtTelefonoMovilReferenciaPersonal2.IsEnabled = true;
            txtTelefonoTrabajoReferenciaPersonal2.IsEnabled = true;
            txtNombreReferenciaPersonal3.IsEnabled = true;
            cmbRelacion3.IsEnabled = true;
            txtTelefonoCasaReferenciaPersonal3.IsEnabled = true;
            txtTelefonoMovilReferenciaPersonal3.IsEnabled = true;
            txtTelefonoTrabajoReferenciaPersonal3.IsEnabled = true;
            txtNombreReferenciaFamiliar1.IsEnabled = true;
            cmbParentesco1.IsEnabled = true;
            txtTelefonoCasaReferenciaFamiliar1.IsEnabled = true;
            txtTelefonoMovilReferenciaFamiliar1.IsEnabled = true;
            txtTelefonoTrabajoReferenciaFamiliar1.IsEnabled = true;
            txtNombreReferenciaFamiliar2.IsEnabled = true;
            cmbParentesco2.IsEnabled = true;
            txtTelefonoCasaReferenciaFamiliar2.IsEnabled = true;
            txtTelefonoMovilReferenciaFamiliar2.IsEnabled = true;
            txtTelefonoTrabajoReferenciaFamiliar2.IsEnabled = true;
            txtNombreReferenciaFamiliar3.IsEnabled = true;
            cmbParentesco3.IsEnabled = true;
            txtTelefonoCasaReferenciaFamiliar3.IsEnabled = true;
            txtTelefonoMovilReferenciaFamiliar3.IsEnabled = true;
            txtTelefonoTrabajoReferenciaFamiliar3.IsEnabled = true;
            txtArticulo.IsEnabled = true;
            txtPrecio.IsEnabled = true;
            txtPlazo.IsEnabled = true;
            txtEnganche.IsEnabled = true;
            txtValorCuota.IsEnabled = true;
            txtSaldo.IsEnabled = true;
            cmbTipoCredito.IsEnabled = true;
            txtComentarios.IsEnabled = true;
        }

        // Metodo para bloquear controles
        private void Bloquear_Controles()
        {
            btnAccion.IsEnabled = false;
            txtSolicitud.IsEnabled = false;
            txtEstadoSolicitud.IsEnabled = false;
            cmbMarca.IsEnabled = false;
            cmbTienda.IsEnabled = false;
            cmbVendedor.IsEnabled = false;
            txtTelefonoVendedor.IsEnabled = false;
            cmbTipoCliente.IsEnabled = false;
            txtFechaIngreso.IsEnabled = false;
            txtPrimerNombreCliente.IsEnabled = false;
            txtSegundoNombreCliente.IsEnabled = false;
            txtPrimerApellidoCliente.IsEnabled = false;
            txtSegundoApellidoCliente.IsEnabled = false;
            txtApellidoCasadaCliente.IsEnabled = false;
            txtDpi.IsEnabled = false;
            txtNacionalidad.IsEnabled = false;
            txtNit.IsEnabled = false;
            txtFechaNacimiento.IsEnabled = false;
            txtEstadoCivil.IsEnabled = false;
            txtSexo.IsEnabled = false;
            txtDependientes.IsEnabled = false;
            txtProfesion.IsEnabled = false;
            txtDireccion.IsEnabled = false;
            txtColonia.IsEnabled = false;
            cmbDepartamento.IsEnabled = false;
            txtMunicipio.IsEnabled = false;
            txtZona.IsEnabled = false;
            cmbTipoVivienda.IsEnabled = false;
            txtTelefonoCasa.IsEnabled = false;
            txtTelefonoMovil.IsEnabled = false;
            txtCorreoElectronico.IsEnabled = false;
            txtNombreEmpresaCliente.IsEnabled = false;
            txtCargo.IsEnabled = false;
            txtFechaIngreso.IsEnabled = false;
            txtIngresosFijos.IsEnabled = false;
            txtIngresosVariables.IsEnabled = false;
            txtOtrosIngresos.IsEnabled = false;
            txtFuenteOtrosIngresos.IsEnabled = false;
            txtDireccionTrabajoCliente.IsEnabled = false;
            txtColoniaTrabajoCliente.IsEnabled = false;
            cmbDepartamentoTrabajoCliente.IsEnabled = false;
            txtMunicipioTrabajoCliente.IsEnabled = false;
            txtZonaTrabajoCliente.IsEnabled = false;
            txtJefeInmediato.IsEnabled = false;
            txtTelefonoTrabajo.IsEnabled = false;
            txtTelefonoJefe.IsEnabled = false;
            txtExtension.IsEnabled = false;
            txtNitEmpresa.IsEnabled = false;
            txtTotalIngresos.IsEnabled = false;
            txtNombresApellidosConyuge.IsEnabled = false;
            txtDpiConyuge.IsEnabled = false;
            txtFechaNacimientoConyuge.IsEnabled = false;
            txtTrabajoConyuge.IsEnabled = false;
            txtCargoConyuge.IsEnabled = false;
            txtFechaIngresoConyuge.IsEnabled = false;
            txtSalarioConyuge.IsEnabled = false;
            txtDireccionTrabajoConyuge.IsEnabled = false;
            txtColoniaTrabajoConyuge.IsEnabled = false;
            cmbDepartamentoTrabajoConyuge.IsEnabled = false;
            txtMunicipioTrabajoConyuge.IsEnabled = false;
            txtZonaTrabajoConyuge.IsEnabled = false;
            txtJefeInmediatoConyuge.IsEnabled = false;
            txtTelefonoTrabajoConyuge.IsEnabled = false;
            txtTelefonoOficinaConyuge.IsEnabled = false;
            txtExtensionConyuge.IsEnabled = false;
            txtTelefonoMovilConyuge.IsEnabled = false;
            txtTotalIngresosConyuge.IsEnabled = false;
            txtNombreReferenciaPersonal1.IsEnabled = false;
            cmbRelacion1.IsEnabled = false;
            txtTelefonoCasaReferenciaPersonal1.IsEnabled = false;
            txtTelefonoMovilReferenciaPersonal1.IsEnabled = false;
            txtTelefonoTrabajoReferenciaPersonal1.IsEnabled = false;
            txtNombreReferenciaPersonal2.IsEnabled = false;
            cmbRelacion2.IsEnabled = false;
            txtTelefonoCasaReferenciaPersonal2.IsEnabled = false;
            txtTelefonoMovilReferenciaPersonal2.IsEnabled = false;
            txtTelefonoTrabajoReferenciaPersonal2.IsEnabled = false;
            txtNombreReferenciaPersonal3.IsEnabled = false;
            cmbRelacion3.IsEnabled = false;
            txtTelefonoCasaReferenciaPersonal3.IsEnabled = false;
            txtTelefonoMovilReferenciaPersonal3.IsEnabled = false;
            txtTelefonoTrabajoReferenciaPersonal3.IsEnabled = false;
            txtNombreReferenciaFamiliar1.IsEnabled = false;
            cmbParentesco1.IsEnabled = false;
            txtTelefonoCasaReferenciaFamiliar1.IsEnabled = false;
            txtTelefonoMovilReferenciaFamiliar1.IsEnabled = false;
            txtTelefonoTrabajoReferenciaFamiliar1.IsEnabled = false;
            txtNombreReferenciaFamiliar2.IsEnabled = false;
            cmbParentesco2.IsEnabled = false;
            txtTelefonoCasaReferenciaFamiliar2.IsEnabled = false;
            txtTelefonoMovilReferenciaFamiliar2.IsEnabled = false;
            txtTelefonoTrabajoReferenciaFamiliar2.IsEnabled = false;
            txtNombreReferenciaFamiliar3.IsEnabled = false;
            cmbParentesco3.IsEnabled = false;
            txtTelefonoCasaReferenciaFamiliar3.IsEnabled = false;
            txtTelefonoMovilReferenciaFamiliar3.IsEnabled = false;
            txtTelefonoTrabajoReferenciaFamiliar3.IsEnabled = false;
            txtArticulo.IsEnabled = false;
            txtPrecio.IsEnabled = false;
            txtPlazo.IsEnabled = false;
            txtEnganche.IsEnabled = false;
            txtValorCuota.IsEnabled = false;
            txtSaldo.IsEnabled = false;
            cmbTipoCredito.IsEnabled = false;
            txtComentarios.IsEnabled = false;
            txtDatosSolicitud.Visibility = Visibility.Hidden;
        }

        // Metodo que Bloquea la aplicación
        private void Bloquear_Tabuladores()
        {
            tabSolicitudes.IsEnabled = false;
            tabSolicitud.IsEnabled = false;
            tabSolicitante.IsEnabled = false;
            tabLaboral.IsEnabled = false;
            tabConyuge.IsEnabled = false;
            tabReferenciasPersonales.IsEnabled = false;
            tabReferenciasLaborales.IsEnabled = false;
            tabFinanciamiento.IsEnabled = false;
            tabComentarios.IsEnabled = false;
        }

        // Metodo que Habilita la aplicación
        private void Habilitar_Tabuladores()
        {
            tabSolicitudes.IsEnabled = true;
            tabSolicitud.IsEnabled = true;
            tabSolicitante.IsEnabled = true;
            tabLaboral.IsEnabled = true;
            tabConyuge.IsEnabled = true;
            tabReferenciasPersonales.IsEnabled = true;
            tabReferenciasLaborales.IsEnabled = true;
            tabFinanciamiento.IsEnabled = true;
            tabComentarios.IsEnabled = true;
        }

        // Inserta estados
        private bool Guardar_Estado(int intSolicitud, int intEstado, ref string refError)
        {
            using (SqlConnection connUtil = new SqlConnection(connStringUTILS))
            {
                connUtil.Open();

                SqlCommand command = connUtil.CreateCommand();

                command.Connection = connUtil;

                string strInsert = "insert into si_estados values ("+intSolicitud+","+intEstado+",getdate(),'"+txtUsuario.Text+"','"+txtComentarios.Text+"')";

                command.CommandText = strInsert;

                try
                {
                    return Convert.ToBoolean(command.ExecuteNonQuery());
                }
                catch (Exception e)
                {
                    refError = e.ToString();
                    return false;
                }
            }
        }

        // Modificar Solicitud
        private bool Modificar_Solicitud(ref string refError)
        {
            using (SqlConnection connUtil = new SqlConnection(connStringUTILS))
            {
                connUtil.Open();

                SqlCommand command = connUtil.CreateCommand();

                command.Connection = connUtil;

                string strUpdate = "UPDATE [dbo].[SI_Solicitudes] " +
                                    "SET [MARCA] = '" + cmbMarca.SelectedItem.ToString() + "',[TIENDA] = '" + cmbTienda.SelectedItem.ToString() + "',[NOMBRE_VENDEDOR] = '" + cmbVendedor.SelectedItem.ToString() +
                                    "',[TELEFONO_VENDEDOR] = '" + txtTelefonoVendedor.Text + "',[TIPO_CLIENTE] = '" + cmbTipoCliente.SelectedItem.ToString() + "',[FECHA_SOLICITUD] = " + Convertir_Fecha(txtFechaSolicitud.Text, 1) +
                                    ",[PRIMER_NOMBRE] = '" + txtPrimerNombreCliente.Text + "',[SEGUNDO_NOMBRE] = '" + txtSegundoNombreCliente.Text + "',[PRIMER_APELLIDO] = '" + txtPrimerApellidoCliente.Text +
                                    "',[SEGUNDO_APELLIDO] = '" + txtSegundoApellidoCliente.Text + "',[APELLIDO_CASADA] = '" + txtApellidoCasadaCliente.Text + "',[DPI_PASAPORTE_CLIENTE] = '" + txtDpi.Text +
                                    "',[NACIONALIDAD] = '" + txtNacionalidad.Text + "',[NIT_CLIENTE] = '" + txtNit.Text + "',[FECHA_NACIMIENTO_CLIENTE] = " + Convertir_Fecha(txtFechaNacimiento.Text, 1) +
                                    ",[ESTADO_CIVIL] = '" + txtEstadoCivil.Text + "',[SEXO] = '" + txtSexo.Text + "',[DEPENDIENTE] = " + txtDependientes.Text + ",[PROFESION] = '" + txtProfesion.Text +
                                    "',[DIRECCION_CLIENTE] = '" + txtDireccion.Text + "',[COLONIA_CLIENTE] = '" + txtColonia.Text + "',[DEPARTAMENTO_CLIENTE] = '" + cmbDepartamento.SelectedItem.ToString() +
                                    "',[MUNICIPIO_CLIENTE] = '" + txtMunicipio.Text + "',[ZONA_CLIENTE] = '" + txtZona.Text + "',[TIPO_VIVIENDA] = '" + cmbTipoVivienda.SelectedItem.ToString() +
                                    "',[TELEFONO_CASA] = '" + txtTelefonoCasa.Text + "',[TELEFONO_MOVIL] = '" + txtTelefonoMovil.Text + "',[CORREO_ELECTRONICO] = '" + txtCorreoElectronico.Text +
                                    "',[NOMBRE_EMPRESA] = '" + txtNombreEmpresaCliente.Text + "',[CARGO_CLIENTE] = '" + txtCargo.Text + "',[FECHA_INGRESO_CLIENTE] = " + Convertir_Fecha(txtFechaIngreso.Text, 1) +
                                    ",[INGRESOS_FIJOS] = " + txtIngresosFijos.Text.Replace("Q", "").Replace(",","") + ",[INGRESOS_VARIABLES] = " + txtIngresosVariables.Text.Replace("Q", "").Replace(",", "") + ",[INGRESOS_OTROS] = " + txtOtrosIngresos.Text.Replace("Q", "").Replace(",", "") +
                                    ",[FUENTES_INGRESOS_OTROS] = '" + txtFuenteOtrosIngresos.Text + "',[DIRECCION_EMPRESA] = '" + txtDireccionTrabajoCliente.Text + "',[COLONIA_EMPRESA] = '" + txtColoniaTrabajoCliente.Text +
                                    "',[DEPARTAMENTO_EMPRESA] = '" + cmbDepartamentoTrabajoCliente.SelectedItem.ToString() + "',[MUNICIPIO_EMPRESA] = '" + txtMunicipioTrabajoCliente.Text + "',[ZONA_EMPRESA] = '" + txtZonaTrabajoCliente.Text +
                                    "',[JEFE_INMEDIATO] = '" + txtJefeInmediato.Text + "',[TELEFONO_TRABAJO] = '" + txtTelefonoTrabajo.Text + "',[TELEFONO_OFICINA] = '" + txtTelefonoJefe.Text + "',[TELEFONO_EXTENSION] = '" + txtExtension.Text +
                                    "',[NIT_EMPRESA] = '" + txtNitEmpresa.Text + "',[NOMBRES_APELLIDOS_CONYUGE] = '" + txtNombresApellidosConyuge.Text + "',[DPI_PASAPORTE_CONYUGE] = '" + txtDpiConyuge.Text +
                                    "',[FECHA_NACIMIENTO_CONYUGE] = " + Convertir_Fecha(txtFechaNacimientoConyuge.Text, 1) + ",[TRABAJO_CONYUGE] = '" + txtTrabajoConyuge.Text + "',[CARGO_CONYUGE] = '" + txtCargoConyuge.Text +
                                    "',[FECHA_INGRESO_CONYUGE] = " + Convertir_Fecha(txtFechaIngresoConyuge.Text, 1) + ",[SALARIO_CONYUGE] = " + txtSalarioConyuge.Text.Replace("Q", "").Replace(",", "") + ",[DIRECCION_TRABAJO_CONYUGE] = '" + txtDireccionTrabajoConyuge.Text +
                                    "',[COLONIA_CONYUGE] = '" + txtColoniaTrabajoConyuge.Text + "',[DEPARTAMENTO_CONYUGE] = '" + cmbDepartamento.SelectedItem.ToString() + "',[MUNICIPIO_CONYUGE] = '" + txtMunicipioTrabajoConyuge.Text +
                                    "',[ZONA_CONYUGE] = '" + txtZonaTrabajoConyuge.Text + "',[JEFE_INMEDIATO_CONYUGE] = '" + txtJefeInmediatoConyuge.Text + "',[TELEFONO_TRABAJO_CONYUGE] = '" + txtTelefonoTrabajoConyuge.Text +
                                    "',[TELEFONO_OFICINA_CONYUGE] = '" + txtTelefonoOficinaConyuge.Text + "',[TELEFONO_EXTENSION_CONYUGE] = '" + txtExtensionConyuge.Text + "',[TELEFONO_MOVIL_CONYUGE] = '" + txtTelefonoMovilConyuge.Text +
                                    "',[TOTAL_INGRESOS_CONYUGE] = " + txtTotalIngresosConyuge.Text.Replace("Q", "") + ",[NOMBRE_REFERENCIA_PERSONAL1] = '" + txtNombreReferenciaPersonal1.Text +
                                    "',[PARENTESCO_REFERENCIA_PERSONAL1] = '" + cmbRelacion1.SelectedItem.ToString() + "',[TELEFONO_CASA_REFERENCIA_PERSONAL1] = '" + txtTelefonoCasaReferenciaPersonal1.Text +
                                    "',[TELEFONO_MOVIL_REFERENCIA_PERSONAL1] = '" + txtTelefonoMovilReferenciaPersonal1.Text + "',[TELEFONO_TRABAJO_REFERENCIA_PERSONAL1] = '" + txtTelefonoTrabajoReferenciaPersonal1.Text +
                                    "',[NOMBRE_REFERENCIA_PERSONAL2] = '" + txtNombreReferenciaPersonal2.Text + "',[PARENTESCO_REFERENCIA_PERSONAL2] = '" + cmbRelacion2.SelectedItem.ToString() +
                                    "',[TELEFONO_CASA_REFERENCIA_PERSONAL2] = '" + txtTelefonoCasaReferenciaPersonal2.Text + "',[TELEFONO_MOVIL_REFERENCIA_PERSONAL2] = '" + txtTelefonoMovilReferenciaPersonal2.Text +
                                    "',[TELEFONO_TRABAJO_REFERENCIA_PERSONAL2] = '" + txtTelefonoTrabajoReferenciaPersonal2.Text + "',[NOMBRE_REFERENCIA_PERSONAL3] = '" + txtNombreReferenciaPersonal3.Text +
                                    "',[PARENTESCO_REFERENCIA_PERSONAL3] = '" + cmbRelacion3.SelectedItem.ToString() + "',[TELEFONO_CASA_REFERENCIA_PERSONAL3] = '" + txtTelefonoCasaReferenciaPersonal3.Text +
                                    "',[TELEFONO_MOVIL_REFERENCIA_PERSONAL3] = '" + txtTelefonoMovilReferenciaPersonal3.Text + "',[TELEFONO_TRABAJO_REFERENCIA_PERSONAL3] = '" + txtTelefonoTrabajoReferenciaPersonal3.Text +
                                    "',[NOMBRE_REFERENCIA_FAMILIAR1] = '" + txtNombreReferenciaFamiliar1.Text + "',[PARENTESCO_REFERENCIA_FAMILIAR1] = '" + cmbParentesco1.SelectedItem.ToString() +
                                    "',[TELEFONO_CASA_REFERENCIA_FAMILIAR1] = '" + txtTelefonoCasaReferenciaFamiliar1.Text + "',[TELEFONO_MOVIL_REFERENCIA_FAMILIAR1] = '" + txtTelefonoMovilReferenciaFamiliar1.Text +
                                    "',[TELEFONO_TRABAJO_REFERENCIA_FAMILIAR1] = '" + txtTelefonoTrabajoReferenciaFamiliar1.Text + "',[NOMBRE_REFERENCIA_FAMILIAR2] = '" + txtNombreReferenciaFamiliar2.Text +
                                    "',[PARENTESCO_REFERENCIA_FAMILIAR2] = '" + cmbParentesco2.SelectedItem.ToString() + "',[TELEFONO_CASA_REFERENCIA_FAMILIAR2] = '" + txtTelefonoCasaReferenciaFamiliar2.Text +
                                    "',[TELEFONO_MOVIL_REFERENCIA_FAMILIAR2] = '" + txtTelefonoMovilReferenciaFamiliar2.Text + "',[TELEFONO_TRABAJO_REFERENCIA_FAMILIAR2] = '" + txtTelefonoTrabajoReferenciaFamiliar2.Text +
                                    "',[NOMBRE_REFERENCIA_FAMILIAR3] = '" + txtNombreReferenciaFamiliar3.Text + "',[PARENTESCO_REFERENCIA_FAMILIAR3] = '" + cmbParentesco3.SelectedItem.ToString() +
                                    "',[TELEFONO_CASA_REFERENCIA_FAMILIAR3] = '" + txtTelefonoCasaReferenciaFamiliar3.Text + "',[TELEFONO_MOVIL_REFERENCIA_FAMILIAR3] = '" + txtTelefonoMovilReferenciaFamiliar3.Text +
                                    "',[TELEFONO_TRABAJO_REFERENCIA_FAMILIAR3] = '" + txtTelefonoTrabajoReferenciaFamiliar3.Text + "',[ARTICULO] = '" + txtArticulo.Text + "',[PRECIO] = " + txtPrecio.Text.Replace("Q", "").Replace(",", "") +
                                    ",[ENGANCHE] = " + txtEnganche.Text.Replace("Q", "").Replace(",", "") + ",[SALDO] = " + txtSaldo.Text.Replace("Q", "").Replace(",", "") + ",[PLAZO] = " + txtPlazo.Text + ",[VALOR_CUOTA] = " + txtValorCuota.Text.Replace("Q", "").Replace(",", "") +
                                    ",[TIPO_CREDITO] = '" + cmbTipoCredito.SelectedItem.ToString() + "',[USUARIO] = '" + txtUsuario.Text + "' WHERE id = " + txtSolicitud.Text;

                command.CommandText = strUpdate;

                try
                {
                    return Convert.ToBoolean(command.ExecuteNonQuery());
                }
                catch (Exception e)
                {
                    refError = e.ToString();
                    return false;
                }
            }
        }

        // Guardar solicitud
        private bool Guardar_Solicitud(ref string refError)
        {
            using (SqlConnection connUtil = new SqlConnection(connStringUTILS))
            {
                connUtil.Open();

                SqlCommand command = connUtil.CreateCommand();

                command.Connection = connUtil;

                string strInsert = "insert into si_solicitudes values('"+Convertir_Combo(cmbMarca.SelectedItem.ToString())+"','"+ Convertir_Combo(cmbTienda.SelectedItem.ToString())+"','"+Vendedor+
                    "','"+txtTelefonoVendedor.Text+"','"+ Convertir_Combo(cmbTipoCliente.SelectedItem.ToString())+"',"+Convertir_Fecha(txtFechaSolicitud.Text,1)+",'"+txtPrimerNombreCliente.Text+
                    "','"+txtSegundoNombreCliente.Text+"','"+txtPrimerApellidoCliente.Text+"','"+txtSegundoApellidoCliente.Text+"','"+txtApellidoCasadaCliente.Text+"','"+txtDpi.Text+
                    "','"+txtNacionalidad.Text+"','"+txtNit.Text+"',"+Convertir_Fecha(txtFechaNacimiento.Text,1)+",'"+txtEstadoCivil.Text+"','"+txtSexo.Text+"','"+txtDependientes.Text+
                    "','"+txtProfesion.Text+"','"+txtDireccion.Text+"','"+txtColonia.Text+"','"+ Convertir_Combo(cmbDepartamento.SelectedItem.ToString())+"','"+txtMunicipio.Text+"','"+txtZona.Text+
                    "','"+ Convertir_Combo(cmbTipoVivienda.SelectedItem.ToString())+"','"+txtTelefonoCasa.Text+"','"+txtTelefonoMovil.Text+"','"+txtCorreoElectronico.Text+"','"+txtNombreEmpresaCliente.Text+
                    "','"+txtCargo.Text+"',"+Convertir_Fecha(txtFechaIngreso.Text,1)+","+Convertir_Moneda(txtIngresosFijos.Text)+","+ Convertir_Moneda(txtIngresosVariables.Text)+","+ Convertir_Moneda(txtOtrosIngresos.Text)+",'"+txtFuenteOtrosIngresos.Text+
                    "','"+txtDireccionTrabajoCliente.Text+"','"+txtColoniaTrabajoCliente.Text+"','"+ Convertir_Combo(cmbDepartamentoTrabajoCliente.SelectedItem.ToString())+"','"+txtMunicipioTrabajoCliente.Text+
                    "','"+txtZonaTrabajoCliente.Text+"','"+txtJefeInmediato.Text+"','"+txtTelefonoTrabajo.Text+"','"+txtTelefonoJefe.Text+"','"+txtExtension.Text+"','"+txtNitEmpresa.Text+
                    "','"+txtNombresApellidosConyuge.Text+"','"+txtDpiConyuge.Text+"',"+Convertir_Fecha(txtFechaNacimientoConyuge.Text,1)+",'"+txtTrabajoConyuge.Text+"','"+txtCargoConyuge.Text+
                    "',"+Convertir_Fecha(txtFechaIngresoConyuge.Text,1)+","+ Convertir_Moneda(txtSalarioConyuge.Text)+",'"+txtDireccionTrabajoConyuge.Text+"','"+txtColoniaTrabajoConyuge.Text+"','"+
                    Convertir_Combo(cmbDepartamentoTrabajoConyuge.SelectedItem.ToString())+"','"+txtMunicipioTrabajoConyuge.Text+"','"+txtZonaTrabajoConyuge.Text+"','"+txtJefeInmediatoConyuge.Text+"','"+
                    txtTelefonoTrabajoConyuge.Text+"','"+txtTelefonoOficinaConyuge.Text+"','"+txtExtensionConyuge.Text+"','"+txtTelefonoMovilConyuge.Text+"',"+ Convertir_Moneda(txtTotalIngresosConyuge.Text)+",'"+
                    txtNombreReferenciaPersonal1.Text+"','"+ Convertir_Combo(cmbRelacion1.SelectedItem.ToString())+"','"+txtTelefonoCasaReferenciaPersonal1.Text+"','"+txtTelefonoMovilReferenciaPersonal1.Text+"','"+
                    txtTelefonoTrabajoReferenciaPersonal1.Text+"','"+txtNombreReferenciaPersonal2.Text+"','"+ Convertir_Combo(cmbRelacion2.SelectedItem.ToString())+"','"+txtTelefonoCasaReferenciaPersonal2.Text+"','"+
                    txtTelefonoMovilReferenciaPersonal2.Text+"','"+txtTelefonoTrabajoReferenciaPersonal2.Text+"','"+txtNombreReferenciaPersonal3.Text+"','"+ Convertir_Combo(cmbRelacion3.SelectedItem.ToString())+"','"+
                    txtTelefonoCasaReferenciaPersonal3.Text+"','"+txtTelefonoMovilReferenciaPersonal3.Text+"','"+txtTelefonoTrabajoReferenciaPersonal3.Text+"','"+txtNombreReferenciaFamiliar1.Text+"','"+
                    Convertir_Combo(cmbParentesco1.SelectedItem.ToString())+"','"+txtTelefonoCasaReferenciaFamiliar1.Text+"','"+txtTelefonoMovilReferenciaFamiliar1.Text+"','"+txtTelefonoTrabajoReferenciaFamiliar1.Text+
                    "','"+txtNombreReferenciaFamiliar2.Text+"','"+cmbParentesco2.Text+"','"+txtTelefonoCasaReferenciaFamiliar2.Text+"','"+txtTelefonoMovilReferenciaFamiliar2.Text+
                    "','"+txtTelefonoTrabajoReferenciaFamiliar2.Text+"','"+txtNombreReferenciaFamiliar3.Text+"','"+ Convertir_Combo(cmbParentesco3.SelectedItem.ToString())+"','"+txtTelefonoCasaReferenciaFamiliar3.Text+
                    "','"+txtTelefonoMovilReferenciaFamiliar3.Text+"','"+txtTelefonoTrabajoReferenciaFamiliar3.Text+"','"+txtArticulo.Text+"',"+ Convertir_Moneda(txtPrecio.Text)+","+ Convertir_Moneda(txtEnganche.Text)+","+ Convertir_Moneda(txtSaldo.Text)+
                    ","+txtPlazo.Text+","+ Convertir_Moneda(txtValorCuota.Text)+",'"+ Convertir_Combo(cmbTipoCredito.SelectedItem.ToString())+"','"+txtUsuario.Text+"')";

                command.CommandText = strInsert;

                try
                {
                    return Convert.ToBoolean(command.ExecuteNonQuery());
                }
                catch (Exception e)
                {
                    refError = e.ToString();
                    return false;
                }
            }
        }

        // Llena el combos con las listas de valores definidas
        private void Llenar_Combos()
        {
            // Llena el combo de marcas
            for (int iMarca = 1; iMarca <= lstMarcas.Count(); iMarca++)
            {
                cmbMarca.Items.Insert(iMarca-1, lstMarcas[iMarca-1].ToUpper());
            }

            // Llena el combo de tiendas
            for (int iTienda = 1; iTienda <= lstTiendas.Count(); iTienda++)
            {
                cmbTienda.Items.Insert(iTienda-1, lstTiendas[iTienda-1].ToUpper());
            }

            //// Llena el combo de vendedores
            //for (int iVendedor = 1; iVendedor <= lstVendedores.Count(); iVendedor++)
            //{
            //    cmbVendedor.Items.Insert(iVendedor-1, lstVendedores[iVendedor-1].ToUpper());
            //}

            // Llena el combo de tipo de clientes
            for (int iTipoCliente = 1; iTipoCliente <= lstTipoClientes.Count(); iTipoCliente++)
            {
                cmbTipoCliente.Items.Insert(iTipoCliente-1, lstTipoClientes[iTipoCliente-1].ToUpper());
            }

            // Llena el combo de tipo de viviendas
            for (int iTipoVivienda = 1; iTipoVivienda <= lstTipoViviendas.Count(); iTipoVivienda++)
            {
                cmbTipoVivienda.Items.Insert(iTipoVivienda-1, lstTipoViviendas[iTipoVivienda-1].ToUpper());
            }

            // Llena el combo de tipo de creditos
            for (int iTipoCredito = 1; iTipoCredito <= lstTipoCreditos.Count(); iTipoCredito++)
            {
                cmbTipoCredito.Items.Insert(iTipoCredito-1, lstTipoCreditos[iTipoCredito-1].ToUpper());
            }

            // Llena el combo de relaciones 1
            for (int iRelacion1 = 1; iRelacion1 <= lstRelaciones.Count(); iRelacion1++)
            {
                cmbRelacion1.Items.Insert(iRelacion1-1, lstRelaciones[iRelacion1-1].ToUpper());
            }

            // Llena el combo de relaciones 2
            for (int iRelacion2 = 1; iRelacion2 <= lstRelaciones.Count(); iRelacion2++)
            {
                cmbRelacion2.Items.Insert(iRelacion2-1, lstRelaciones[iRelacion2-1].ToUpper());
            }

            // Llena el combo de relaciones 3
            for (int iRelacion3 = 1; iRelacion3 <= lstRelaciones.Count(); iRelacion3++)
            {
                cmbRelacion3.Items.Insert(iRelacion3-1, lstRelaciones[iRelacion3-1].ToUpper());
            }

            // Llena el combo de parentesco 1 
            for (int iParentesco1 = 1; iParentesco1 <= lstParentescos.Count(); iParentesco1++)
            {
                cmbParentesco1.Items.Insert(iParentesco1-1, lstParentescos[iParentesco1-1].ToUpper());
            }

            // Llena el combo de parentesco 2 
            for (int iParentesco2 = 1; iParentesco2 <= lstParentescos.Count(); iParentesco2++)
            {
                cmbParentesco2.Items.Insert(iParentesco2-1, lstParentescos[iParentesco2-1].ToUpper());
            }

            // Llena el combo de parentesco 3 
            for (int iParentesco3 = 1; iParentesco3 <= lstParentescos.Count(); iParentesco3++)
            {
                cmbParentesco3.Items.Insert(iParentesco3-1, lstParentescos[iParentesco3-1].ToUpper());
            }

            // Llena el combo de departamento
            for (int iDepartamento = 1; iDepartamento <= lstDepartamentos.Count(); iDepartamento++)
            {
                cmbDepartamento.Items.Insert(iDepartamento-1, lstDepartamentos[iDepartamento-1].ToUpper());
            }

            // Llena el combo de departamento trabajo cliente
            for (int iDepartamentoTrabajoCliente = 1; iDepartamentoTrabajoCliente <= lstDepartamentos.Count(); iDepartamentoTrabajoCliente++)
            {
                cmbDepartamentoTrabajoCliente.Items.Insert(iDepartamentoTrabajoCliente-1, lstDepartamentos[iDepartamentoTrabajoCliente-1].ToUpper());
            }

            // Llena el combo de departamento trabajo conyuge
            for (int iDepartamentoTrabajoConyuge = 1; iDepartamentoTrabajoConyuge <= lstDepartamentos.Count(); iDepartamentoTrabajoConyuge++)
            {
                cmbDepartamentoTrabajoConyuge.Items.Insert(iDepartamentoTrabajoConyuge-1, lstDepartamentos[iDepartamentoTrabajoConyuge-1].ToUpper());
            }

            // Llena el combo de acciones del analista de credito
            for (int iAcciones = 1; iAcciones <= lstAcciones.Count(); iAcciones++)
            {
                cmbAccion.Items.Insert(iAcciones, lstAcciones[iAcciones - 1].ToUpper());
            }
        }
    }
}
