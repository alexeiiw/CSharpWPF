   M S B A M L   `   `   `  ���� T  PMahApps.Metro, Version=2.0.0.0, Culture=neutral, PublicKeyToken=51482d6f650b2b3fV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls  G CScoreInterno, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:ScoreInternoScoreInterno    ScoreInterno.MainWindow   Cmah;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro  \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     ,localclr-namespace:ScoreInterno     -     � Title$!  SCORE INTERNO V1 Desarrollo��$	��475��$	��800��(   "MahApps.Metro.Controls.MetroWindow   ShowMaxRestoreButton$ False��   RightWindowCommands +   %MahApps.Metro.Controls.WindowCommands  .@��� -   $��btnABC��$��Nuevo��${�False���� -   $��btnCancelar��$��Cancelar��${�False��.��� .+��� -    ��Name$ 
tabControl��.@��� -   $��tabSolicitudes��$��*��.���� -    ��Name$ ScrollSolicitudes���� -   $��tabSolicitud��$��#��.��� ���	#FFE5E5E5.+��� -    ��Name$ label8��$��DATOS DE LA SOLICITUD��$��Center=�$��	10,10,0,0q�$	��Top=�$��16(�
����  �$
��Bold"��� -	   $ lblSolicitud��$��No. Solicitud:��$
��Left=�$��	10,50,0,0q�$	��Top=�$
��Bold"��� -
    ��Name$ txtSolicitud��$
��Left=�$��
110,50,0,0q�$�� ��$	��Top=�$	��100��$��0�� �� CharacterCasing$ Upper=��� -   $ 	lblAccion��$��Acción:��$
��Left=�$��	10,50,0,0q�$	��Top=�$
��Bold"�$y�Hidden=��� -    ��Name$ 	cmbAccion��$��	75,50,0,0q�$	��Top=�$	��100��$
��Left=�$y�Hidden=�.@��� 	 �� 
IsSelected$
	 True��.��
ACCIONES�� -   
 ��Name$
 	btnAccion��$��Acción��$
��Left=�$��
230,50,0,0q�$	��Top=�$��1���� -   $ label6��$��Estado:��$
��Left=�$��
370,50,0,0q�$	��Top=�$
��Bold"��� -   $ txtEstadoSolicitud��$
��Left=�$��
430,50,0,0q�$�� ��$	��Top=�$	��150��$ Upper=��� -   $ txtDatosSolicitud��$
��Left=�$��
590,50,0,0q�$�� ��$	��Top=�$	��100��$ Upper=�$y�Hidden=�${�False���� -   $ label1��$��Marca:��$
��Left=�$��	10,90,0,0q�$	��Top=�$
��Bold"��� -   $ cmbMarca��$��	75,90,0,0q�$	��Top=�$	��150��$
��Left=�$��2��.@��� $
	 True��.��MARCAS�� -   $ label2��$��Tienda:��$
��Left=�$��
250,90,0,0q�$	��Top=�$
��Bold"��� -   $ 	cmbTienda��$��
335,90,0,0q�$	��Top=�$	��150��$
��Left=�$��3��.@��� $
	 True��.��	TIENDAS�� -   $ lblVendedor��$��	Vendedor:��$
��Left=�$��
505,90,0,0q�$	��Top=�$
��Bold"��� -   $ cmbVendedor��$��
590,90,0,0q�$	��Top=�$	��150��$
��Left=�$��4��.@��� $
	 True��.��
VENDEDORES�� -   $ label4��$��
Teléfono:��$
��Left=�$��
10,130,0,0q�$	��Top=�$
��Bold"��� -   $ txtTelefonoVendedor��$
��Left=�$��
75,130,0,0q�$�� ��$	��Top=�$	��150��$��5��$ Upper=��� -   $ label7��$��Tipo Cliente:��$
��Left=�$��250,130,0,0q�$	��Top=�$
��Bold"��� -   $ cmbTipoCliente��$��335,130,0,0q�$	��Top=�$	��150��$
��Left=�$��6��.@��� $
	 True��.��TIPO DE CLIENTES�� -   $ label5��$��Fecha:��$
��Left=�$��505,130,0,0q�$	��Top=�$
��Bold"��� -   $ txtFechaSolicitud��$
��Left=�$��590,130,0,0q�$�� ��$	��Top=�$	��150��$��7��$ Upper=��� -   $ label17��$��Comentarios:��$
��Left=�$��
10,170,0,0q�$	��Top=�$
��Bold"��� -   $ txtComentarios��$
��Left=�$��
10,200,0,0q�$�� ��$	��Top=�$	��730��$��8��$ Upper=�$	��100���� -   $ lblMensajes��$��Mensaje��$
��Left=�$��
10,360,0,0q�$	��Top=�$
��Bold"�$y�Hidden=�$
��Auto=�
���  ���� -    $��tabSolicitante��$��Solicitante��.��� ���	#FFE5E5E5.+��� -!   $ label9��$��DATOS DEL SOLICITANTE��$��Center=�$��	10,10,0,0q�$	��Top=�$��16(�
����  �$
��Bold"��� -"   $ txtPrimerNombreCliente��$	��180��$	��Top=�$
��Left=�$��	10,50,0,0q�$ Upper=���  J� IsHitTestVisible$ False��$��PRIMER NOMBRE��$	��Top=�$
��Left=�$��	15,54,0,0q�
���  ������  �� 
TargetType" ��.���  �� Property � y� �� Value$ 	Collapsed=� �� Triggers ��  �� Binding �� *Text+ �� ElementName$ txtPrimerNombreCliente�� �� Value  .a���  � y�$ Visible=��� -#   $ txtSegundoNombreCliente��$	��180��$	��Top=�$
��Left=�$��
200,50,0,0q�$ Upper=��� $ False��$��SEGUNDO NOMBRE��$	��Top=�$
��Left=�$��
205,54,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtSegundoNombreCliente��  .a���  � y�$ Visible=��� -$   $ txtPrimerApellidoCliente��$	��180��$	��Top=�$
��Left=�$��
390,50,0,0q�$ Upper=��� $ False��$��PRIMER APELLIDO��$	��Top=�$
��Left=�$��
395,54,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtPrimerApellidoCliente��  .a���  � y�$ Visible=��� -%   $ txtSegundoApellidoCliente��$	��180��$	��Top=�$
��Left=�$��
580,50,0,0q�$ Upper=��� $ False��$��SEGUNDO APELLIDO��$	��Top=�$
��Left=�$��
585,54,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtSegundoApellidoCliente��  .a���  � y�$ Visible=��� -&   $ txtApellidoCasadaCliente��$	��180��$	��Top=�$
��Left=�$��	10,90,0,0q�$ Upper=��� $ False��$��APELLIDO CASADA��$	��Top=�$
��Left=�$��	15,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtApellidoCasadaCliente��  .a���  � y�$ Visible=��� -'   $ txtDpi��$	��180��$	��Top=�$
��Left=�$��
200,90,0,0q�$ Upper=� � ToolTip$ Ejemplo: 2421856580101���� $ False��$��DPI PASAPORTE��$	��Top=�$
��Left=�$��
205,94,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtDpi��  .a���  � y�$ Visible=��� -(   $ txtNacionalidad��$	��180��$	��Top=�$
��Left=�$��
390,90,0,0q�$ Upper=��� $ False��$��NACIONALIDAD��$	��Top=�$
��Left=�$��
395,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtNacionalidad��  .a���  � y�$ Visible=��� -)   $ txtNit��$	��180��$	��Top=�$
��Left=�$��
580,90,0,0q�$ Upper=��� $ False��$	��NIT��$	��Top=�$
��Left=�$��
585,94,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtNit��  .a���  � y�$ Visible=��� -*   $ txtFechaNacimiento��$	��180��$	��Top=�$
��Left=�$��
10,130,0,0q�$ Upper=��� $ False��$��FECHA NACIMIENTO��$	��Top=�$
��Left=�$��
15,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtFechaNacimiento��  .a���  � y�$ Visible=��� -+   $ txtEstadoCivil��$	��180��$	��Top=�$
��Left=�$��200,130,0,0q�$ Upper=��� $ False��$��ESTADO CIVIL��$	��Top=�$
��Left=�$��205,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtEstadoCivil��  .a���  � y�$ Visible=��� -,   $ txtSexo��$	��180��$	��Top=�$
��Left=�$��390,130,0,0q�$ Upper=��� $ False��$
��SEXO��$	��Top=�$
��Left=�$��395,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtSexo��  .a���  � y�$ Visible=��� --   $ txtDependientes��$	��180��$	��Top=�$
��Left=�$��580,130,0,0q�$ Upper=�$��Right=��� $ False��$��DEPENDIENTES��$	��Top=�$
��Left=�$��585,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtDependientes��  .a���  � y�$ Visible=��� -.   $ txtProfesion��$	��180��$	��Top=�$
��Left=�$��
10,170,0,0q�$ Upper=��� $ False��$��
PROFESIÓN��$	��Top=�$
��Left=�$��
15,174,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtProfesion��  .a���  � y�$ Visible=��� -/   $ txtDireccion��$	��370��$	��Top=�$
��Left=�$��200,170,0,0q�$ Upper=��� $ False��$��DIRECCIÓN DOMICILIO��$	��Top=�$
��Left=�$��205,174,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtDireccion��  .a���  � y�$ Visible=��� -0   $ 
txtColonia��$	��180��$	��Top=�$
��Left=�$��580,170,0,0q�$ Upper=��� $ False��$��COLONIA ALDEA CASERIO��$	��Top=�$
��Left=�$��585,174,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ 
txtColonia��  .a���  � y�$ Visible=��� -1   $ cmbDepartamento��$��
10,210,0,0q�$	��Top=�$	��180��$
��Left=�.@��� $
	 True��.��DEPARTAMENTOS�� -2   $ txtMunicipio��$	��180��$	��Top=�$
��Left=�$��200,210,0,0q�$ Upper=��� $ False��$��	MUNICIPIO��$	��Top=�$
��Left=�$��205,214,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtMunicipio��  .a���  � y�$ Visible=��� -3   $ txtZona��$	��180��$	��Top=�$
��Left=�$��390,210,0,0q�$ Upper=�$��Right=��� $ False��$
��ZONA��$	��Top=�$
��Left=�$��395,214,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtZona��  .a���  � y�$ Visible=��� -4   $ cmbTipoVivienda��$��580,210,0,0q�$	��Top=�$	��180��$
��Left=�.@��� $
	 True��.��TIPO DE VIVIENDAS�� -5   $ txtTelefonoCasa��$	��180��$	��Top=�$
��Left=�$��
10,250,0,0q�$ Upper=��� $ False��$��TELÉFONO CASA��$	��Top=�$
��Left=�$��
15,254,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtTelefonoCasa��  .a���  � y�$ Visible=��� -6   $ txtTelefonoMovil��$	��180��$	��Top=�$
��Left=�$��200,250,0,0q�$ Upper=��� $ False��$��TELÉFONO MÓVIL��$	��Top=�$
��Left=�$��205,254,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtTelefonoMovil��  .a���  � y�$ Visible=��� -7   $ txtCorreoElectronico��$	��180��$	��Top=�$
��Left=�$��390,250,0,0q�$ Upper=��� $ False��$��CORREO ELECTRÓNICO��$	��Top=�$
��Left=�$��395,254,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtCorreoElectronico��  .a���  � y�$ Visible=��� -8   $��
tabLaboral��$��Laboral��.��� ���	#FFE5E5E5.+��� -9   $ label10��$%��DATOS LABORALES DEL SOLICITANTE��$��Center=�$��	10,10,0,0q�$	��Top=�$��16(�
����  �$
��Bold"��� -:   $ txtNombreEmpresaCliente��$	��370��$	��Top=�$
��Left=�$��	10,50,0,0q�$ Upper=��� $ False��$��NOMBRE EMPRESA NEGOCIO��$	��Top=�$
��Left=�$��	15,54,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtNombreEmpresaCliente��  .a���  � y�$ Visible=��� -;   $ txtCargo��$	��180��$	��Top=�$
��Left=�$��
390,50,0,0q�$ Upper=��� $ False��$��CARGO PUESTO��$	��Top=�$
��Left=�$��
395,54,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtCargo��  .a���  � y�$ Visible=��� -<   $ txtFechaIngreso��$	��180��$	��Top=�$
��Left=�$��
580,50,0,0q�$ Upper=��� $ False��$��FECHA INGRESO��$	��Top=�$
��Left=�$��
585,54,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtFechaIngreso��  .a���  � y�$ Visible=��� -=   $ txtIngresosFijos��$	��180��$	��Top=�$
��Left=�$��	10,90,0,0q�$ Upper=�$��Right=��� $ False��$��INGRESOS FIJOS Q.��$	��Top=�$
��Left=�$��	15,94,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtIngresosFijos��  .a���  � y�$ Visible=��� ->   $ txtIngresosVariables��$	��180��$	��Top=�$
��Left=�$��
200,90,0,0q�$ Upper=�$��Right=��� $ False��$��INGRESOS VARIABLES Q.��$	��Top=�$
��Left=�$��
205,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtIngresosVariables��  .a���  � y�$ Visible=��� -?   $ txtOtrosIngresos��$	��180��$	��Top=�$
��Left=�$��
390,90,0,0q�$ Upper=�$��Right=��� $ False��$��OTROS INGRESOS Q.��$	��Top=�$
��Left=�$��
395,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtOtrosIngresos��  .a���  � y�$ Visible=��� -@   $ txtFuenteOtrosIngresos��$	��180��$	��Top=�$
��Left=�$��
580,90,0,0q�$ Upper=��� $ False��$��FUENTE OTROS INGRESOS��$	��Top=�$
��Left=�$��
585,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtFuenteOtrosIngresos��  .a���  � y�$ Visible=��� -A   $  txtDireccionTrabajoCliente��$	��370��$	��Top=�$
��Left=�$��
10,130,0,0q�$ Upper=��� $ False��$ ��DIRECCIÓN TRABAJO NEGOCIO��$	��Top=�$
��Left=�$��
15,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$  txtDireccionTrabajoCliente��  .a���  � y�$ Visible=��� -B   $ txtColoniaTrabajoCliente��$	��180��$	��Top=�$
��Left=�$��390,130,0,0q�$ Upper=��� $ False��$��COLONIA ALDEA CASERIO��$	��Top=�$
��Left=�$��395,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtColoniaTrabajoCliente��  .a���  � y�$ Visible=��� -C   $# cmbDepartamentoTrabajoCliente��$��580,130,0,0q�$	��Top=�$	��180��$
��Left=�.@��� $
	 True��.��DEPARTAMENTOS�� -D   $  txtMunicipioTrabajoCliente��$	��180��$	��Top=�$
��Left=�$��
10,170,0,0q�$ Upper=��� $ False��$��	MUNICIPIO��$	��Top=�$
��Left=�$��
15,174,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$  txtMunicipioTrabajoCliente��  .a���  � y�$ Visible=��� -E   $ txtZonaTrabajoCliente��$	��180��$	��Top=�$
��Left=�$��200,170,0,0q�$ Upper=�$��Right=��� $ False��$
��ZONA��$	��Top=�$
��Left=�$��205,174,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtZonaTrabajoCliente��  .a���  � y�$ Visible=��� -F   $ txtJefeInmediato��$	��180��$	��Top=�$
��Left=�$��390,170,0,0q�$ Upper=��� $ False��$��JEFE INMEDIATO��$	��Top=�$
��Left=�$��395,174,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtJefeInmediato��  .a���  � y�$ Visible=��� -G   $ txtTelefonoTrabajo��$	��180��$	��Top=�$
��Left=�$��580,170,0,0q�$ Upper=��� $ False��$��TELÉFONO TRABAJO��$	��Top=�$
��Left=�$��585,174,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtTelefonoTrabajo��  .a���  � y�$ Visible=��� -H   $ txtTelefonoJefe��$	��180��$	��Top=�$
��Left=�$��
10,210,0,0q�$ Upper=��� $ False��$��TELÉFONO JEFE INMEDIATO��$	��Top=�$
��Left=�$��
15,214,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtTelefonoJefe��  .a���  � y�$ Visible=��� -I   $ txtExtension��$	��180��$	��Top=�$
��Left=�$��200,210,0,0q�$ Upper=��� $ False��$��
EXTENSIÓN��$	��Top=�$
��Left=�$��205,214,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtExtension��  .a���  � y�$ Visible=��� -J   $ txtNitEmpresa��$	��180��$	��Top=�$
��Left=�$��390,210,0,0q�$ Upper=��� $ False��$��NIT EMPRESA��$	��Top=�$
��Left=�$��395,214,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtNitEmpresa��  .a���  � y�$ Visible=��� -K   $ txtTotalIngresos��$	��180��$	��Top=�$
��Left=�$��580,210,0,0q�$ Upper=�$��Right=��� $ False��$��TOTAL INGRESOS Q.��$	��Top=�$
��Left=�$��585,214,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtTotalIngresos��  .a���  � y�$ Visible=��� -L   $��
tabConyuge��$��Conyuge��.��� ���	#FFE5E5E5.+��� -M   $ label11��$��DATOS DEL CONYUGE��$��Center=�$��	10,10,0,0q�$	��Top=�$��16(�
����  �$
��Bold"��� -N   $  txtNombresApellidosConyuge��$	��370��$	��Top=�$
��Left=�$��	10,50,0,0q�$ Upper=��� $ False��$��NOMBRES APELLIDOS��$	��Top=�$
��Left=�$��	15,54,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$  txtNombresApellidosConyuge��  .a���  � y�$ Visible=��� -O   $ txtDpiConyuge��$	��180��$	��Top=�$
��Left=�$��
390,50,0,0q�$ Upper=�$ Ejemplo: 2421856580101���� $ False��$��DPI PASAPORTE��$	��Top=�$
��Left=�$��
395,54,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtDpiConyuge��  .a���  � y�$ Visible=��� -P   $ txtFechaNacimientoConyuge��$	��180��$	��Top=�$
��Left=�$��
580,50,0,0q�$ Upper=��� $ False��$��FECHA NACIMIENTO��$	��Top=�$
��Left=�$��
585,54,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtFechaNacimientoConyuge��  .a���  � y�$ Visible=��� -Q   $ txtTrabajoConyuge��$	��370��$	��Top=�$
��Left=�$��	10,90,0,0q�$ Upper=��� $ False��$��TRABAJO NEGOCIO��$	��Top=�$
��Left=�$��	15,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtTrabajoConyuge��  .a���  � y�$ Visible=��� -R   $ txtCargoConyuge��$	��180��$	��Top=�$
��Left=�$��
390,90,0,0q�$ Upper=��� $ False��$��CARGO PUESTO��$	��Top=�$
��Left=�$��
395,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtCargoConyuge��  .a���  � y�$ Visible=��� -S   $ txtFechaIngresoConyuge��$	��180��$	��Top=�$
��Left=�$��
580,90,0,0q�$ Upper=��� $ False��$��FECHA INGRESO��$	��Top=�$
��Left=�$��
585,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtFechaIngresoConyuge��  .a���  � y�$ Visible=��� -T   $ txtSalarioConyuge��$	��180��$	��Top=�$
��Left=�$��
10,130,0,0q�$ Upper=�$��Right=��� $ False��$��
SALARIO Q.��$	��Top=�$
��Left=�$��
15,134,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtSalarioConyuge��  .a���  � y�$ Visible=��� -U   $  txtDireccionTrabajoConyuge��$	��370��$	��Top=�$
��Left=�$��200,130,0,0q�$ Upper=��� $ False��$ ��DIRECCIÓN TRABAJO NEGOCIO��$	��Top=�$
��Left=�$��205,134,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$  txtDireccionTrabajoConyuge��  .a���  � y�$ Visible=��� -V   $ txtColoniaTrabajoConyuge��$	��180��$	��Top=�$
��Left=�$��580,130,0,0q�$ Upper=��� $ False��$��COLONIA ALDEA CASERIO��$	��Top=�$
��Left=�$��585,134,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtColoniaTrabajoConyuge��  .a���  � y�$ Visible=��� -W   $# cmbDepartamentoTrabajoConyuge��$��
10,170,0,0q�$	��Top=�$	��180��$
��Left=�.@��� $
	 True��.��DEPARTAMENTOS�� -X   $  txtMunicipioTrabajoConyuge��$	��180��$	��Top=�$
��Left=�$��200,170,0,0q�$ Upper=��� $ False��$��	MUNICIPIO��$	��Top=�$
��Left=�$��205,174,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$  txtMunicipioTrabajoConyuge��  .a���  � y�$ Visible=��� -Y   $ txtZonaTrabajoConyuge��$	��180��$	��Top=�$
��Left=�$��390,170,0,0q�$ Upper=�$��Right=��� $ False��$
��ZONA��$	��Top=�$
��Left=�$��395,174,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtZonaTrabajoConyuge��  .a���  � y�$ Visible=��� -Z   $ txtJefeInmediatoConyuge��$	��180��$	��Top=�$
��Left=�$��580,170,0,0q�$ Upper=��� $ False��$��JEFE INMEDIATO��$	��Top=�$
��Left=�$��585,174,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtJefeInmediatoConyuge��  .a���  � y�$ Visible=��� -[   $ txtTelefonoTrabajoConyuge��$	��180��$	��Top=�$
��Left=�$��
10,210,0,0q�$ Upper=��� $ False��$��TELÉFONO TRABAJO��$	��Top=�$
��Left=�$��
15,214,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtTelefonoTrabajoConyuge��  .a���  � y�$ Visible=��� -\   $ txtTelefonoOficinaConyuge��$	��180��$	��Top=�$
��Left=�$��200,210,0,0q�$ Upper=��� $ False��$��TELÉFONO OFICINA��$	��Top=�$
��Left=�$��205,214,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtTelefonoOficinaConyuge��  .a���  � y�$ Visible=��� -]   $ txtExtensionConyuge��$	��180��$	��Top=�$
��Left=�$��390,210,0,0q�$ Upper=��� $ False��$��
EXTENSIÓN��$	��Top=�$
��Left=�$��395,214,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtExtensionConyuge��  .a���  � y�$ Visible=��� -^   $ txtTelefonoMovilConyuge��$	��180��$	��Top=�$
��Left=�$��580,210,0,0q�$ Upper=��� $ False��$��TELÉFONO MÓVIL��$	��Top=�$
��Left=�$��585,214,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtTelefonoMovilConyuge��  .a���  � y�$ Visible=��� -_   $ txtTotalIngresosConyuge��$	��180��$	��Top=�$
��Left=�$��
10,250,0,0q�$ Upper=�$��Right=��� $ False��$��TOTAL INGRESOS Q.��$	��Top=�$
��Left=�$��
15,254,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtTotalIngresosConyuge��  .a���  � y�$ Visible=��� -`   $��tabReferenciasPersonales��$��
Personales��.��� ���	#FFE5E5E5.+��� -a   $ label12��$��REFERENCIAS PERSONALES��$��Center=�$��	10,10,0,0q�$	��Top=�$��16(�
����  �$
��Bold"��� -b   $" txtNombreReferenciaPersonal1��$	��180��$	��Top=�$
��Left=�$��	10,50,0,0q�$ Upper=��� $ False��$��NOMBRE��$	��Top=�$
��Left=�$��	15,54,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$" txtNombreReferenciaPersonal1��  .a���  � y�$ Visible=��� -c   $ cmbRelacion1��$��
200,50,0,0q�$	��Top=�$	��135��$
��Left=�.@��� $
	 True��.��
RELACIONES�� -d   $( "txtTelefonoCasaReferenciaPersonal1��$	��135��$	��Top=�$
��Left=�$��
345,50,0,0q�$ Upper=��� $ False��$��	TEL. CASA��$	��Top=�$
��Left=�$��
350,54,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$( "txtTelefonoCasaReferenciaPersonal1��  .a���  � y�$ Visible=��� -e   $) #txtTelefonoMovilReferenciaPersonal1��$	��135��$	��Top=�$
��Left=�$��
490,50,0,0q�$ Upper=��� $ False��$��TEL. MÓVIL��$	��Top=�$
��Left=�$��
495,54,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$) #txtTelefonoMovilReferenciaPersonal1��  .a���  � y�$ Visible=��� -f   $+ %txtTelefonoTrabajoReferenciaPersonal1��$	��135��$	��Top=�$
��Left=�$��
635,50,0,0q�$ Upper=��� $ False��$��TEL. TRABAJO��$	��Top=�$
��Left=�$��
640,54,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$+ %txtTelefonoTrabajoReferenciaPersonal1��  .a���  � y�$ Visible=��� -g   $" txtNombreReferenciaPersonal2��$	��180��$	��Top=�$
��Left=�$��	10,90,0,0q�$ Upper=��� $ False��$��NOMBRE��$	��Top=�$
��Left=�$��	15,94,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$" txtNombreReferenciaPersonal2��  .a���  � y�$ Visible=��� -h   $ cmbRelacion2��$��
200,90,0,0q�$	��Top=�$	��135��$
��Left=�.@��� $
	 True��.��
RELACIONES�� -i   $( "txtTelefonoCasaReferenciaPersonal2��$	��135��$	��Top=�$
��Left=�$��
345,90,0,0q�$ Upper=��� $ False��$��	TEL. CASA��$	��Top=�$
��Left=�$��
350,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$( "txtTelefonoCasaReferenciaPersonal2��  .a���  � y�$ Visible=��� -j   $) #txtTelefonoMovilReferenciaPersonal2��$	��135��$	��Top=�$
��Left=�$��
490,90,0,0q�$ Upper=��� $ False��$��TEL. MÓVIL��$	��Top=�$
��Left=�$��
495,94,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$) #txtTelefonoMovilReferenciaPersonal2��  .a���  � y�$ Visible=��� -k   $+ %txtTelefonoTrabajoReferenciaPersonal2��$	��135��$	��Top=�$
��Left=�$��
635,90,0,0q�$ Upper=��� $ False��$��TEL. TRABAJO��$	��Top=�$
��Left=�$��
640,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$+ %txtTelefonoTrabajoReferenciaPersonal2��  .a���  � y�$ Visible=��� -l   $" txtNombreReferenciaPersonal3��$	��180��$	��Top=�$
��Left=�$��
10,130,0,0q�$ Upper=��� $ False��$��NOMBRE��$	��Top=�$
��Left=�$��
15,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$" txtNombreReferenciaPersonal3��  .a���  � y�$ Visible=��� -m   $ cmbRelacion3��$��200,130,0,0q�$	��Top=�$	��135��$
��Left=�.@��� $
	 True��.��
RELACIONES�� -n   $( "txtTelefonoCasaReferenciaPersonal3��$	��135��$	��Top=�$
��Left=�$��345,130,0,0q�$ Upper=��� $ False��$��	TEL. CASA��$	��Top=�$
��Left=�$��350,134,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$( "txtTelefonoCasaReferenciaPersonal3��  .a���  � y�$ Visible=��� -o   $) #txtTelefonoMovilReferenciaPersonal3��$	��135��$	��Top=�$
��Left=�$��490,130,0,0q�$ Upper=��� $ False��$��TEL. MÓVIL��$	��Top=�$
��Left=�$��495,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$) #txtTelefonoMovilReferenciaPersonal3��  .a���  � y�$ Visible=��� -p   $+ %txtTelefonoTrabajoReferenciaPersonal3��$	��135��$	��Top=�$
��Left=�$��635,130,0,0q�$ Upper=��� $ False��$��TEL. TRABAJO��$	��Top=�$
��Left=�$��640,134,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$+ %txtTelefonoTrabajoReferenciaPersonal3��  .a���  � y�$ Visible=��� -q   $��tabReferenciasLaborales��$��
Familiares��.��� ���	#FFE5E5E5.+��� -r   $ label13��$��REFERENCIAS FAMILIARES��$��Center=�$��	10,10,0,0q�$	��Top=�$��16(�
����  �$
��Bold"��� -s   $" txtNombreReferenciaFamiliar1��$	��180��$	��Top=�$
��Left=�$��	10,50,0,0q�$ Upper=��� $ False��$��NOMBRE��$	��Top=�$
��Left=�$��	15,54,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$" txtNombreReferenciaFamiliar1��  .a���  � y�$ Visible=��� -t   $ cmbParentesco1��$��
200,50,0,0q�$	��Top=�$	��135��$
��Left=�.@��� $
	 True��.��PARENTESCOS�� -u   $( "txtTelefonoCasaReferenciaFamiliar1��$	��135��$	��Top=�$
��Left=�$��
345,50,0,0q�$ Upper=��� $ False��$��	TEL. CASA��$	��Top=�$
��Left=�$��
350,54,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$( "txtTelefonoCasaReferenciaFamiliar1��  .a���  � y�$ Visible=��� -v   $) #txtTelefonoMovilReferenciaFamiliar1��$	��135��$	��Top=�$
��Left=�$��
490,50,0,0q�$ Upper=��� $ False��$��TEL. MÓVIL��$	��Top=�$
��Left=�$��
495,54,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$) #txtTelefonoMovilReferenciaFamiliar1��  .a���  � y�$ Visible=��� -w   $+ %txtTelefonoTrabajoReferenciaFamiliar1��$	��135��$	��Top=�$
��Left=�$��
635,50,0,0q�$ Upper=��� $ False��$��TEL. TRABAJO��$	��Top=�$
��Left=�$��
640,54,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$+ %txtTelefonoTrabajoReferenciaFamiliar1��  .a���  � y�$ Visible=��� -x   $" txtNombreReferenciaFamiliar2��$	��180��$	��Top=�$
��Left=�$��	10,90,0,0q�$ Upper=��� $ False��$��NOMBRE��$	��Top=�$
��Left=�$��	15,94,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$" txtNombreReferenciaFamiliar2��  .a���  � y�$ Visible=��� -y   $ cmbParentesco2��$��
200,90,0,0q�$	��Top=�$	��135��$
��Left=�.@��� $
	 True��.��PARENTESCOS�� -z   $( "txtTelefonoCasaReferenciaFamiliar2��$	��135��$	��Top=�$
��Left=�$��
345,90,0,0q�$ Upper=��� $ False��$��	TEL. CASA��$	��Top=�$
��Left=�$��
350,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$( "txtTelefonoCasaReferenciaFamiliar2��  .a���  � y�$ Visible=��� -{   $) #txtTelefonoMovilReferenciaFamiliar2��$	��135��$	��Top=�$
��Left=�$��
490,90,0,0q�$ Upper=��� $ False��$��TEL. MÓVIL��$	��Top=�$
��Left=�$��
495,94,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$) #txtTelefonoMovilReferenciaFamiliar2��  .a���  � y�$ Visible=��� -|   $+ %txtTelefonoTrabajoReferenciaFamiliar2��$	��135��$	��Top=�$
��Left=�$��
635,90,0,0q�$ Upper=��� $ False��$��TEL. TRABAJO��$	��Top=�$
��Left=�$��
640,94,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$+ %txtTelefonoTrabajoReferenciaFamiliar2��  .a���  � y�$ Visible=��� -}   $" txtNombreReferenciaFamiliar3��$	��180��$	��Top=�$
��Left=�$��
10,130,0,0q�$ Upper=��� $ False��$��NOMBRE��$	��Top=�$
��Left=�$��
15,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$" txtNombreReferenciaFamiliar3��  .a���  � y�$ Visible=��� -~   $ cmbParentesco3��$��200,130,0,0q�$	��Top=�$	��135��$
��Left=�.@��� $
	 True��.��PARENTESCOS�� -   $( "txtTelefonoCasaReferenciaFamiliar3��$	��135��$	��Top=�$
��Left=�$��345,130,0,0q�$ Upper=��� $ False��$��	TEL. CASA��$	��Top=�$
��Left=�$��350,134,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$( "txtTelefonoCasaReferenciaFamiliar3��  .a���  � y�$ Visible=��� -�   $) #txtTelefonoMovilReferenciaFamiliar3��$	��135��$	��Top=�$
��Left=�$��490,130,0,0q�$ Upper=��� $ False��$��TEL. MÓVIL��$	��Top=�$
��Left=�$��495,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$) #txtTelefonoMovilReferenciaFamiliar3��  .a���  � y�$ Visible=��� -�   $+ %txtTelefonoTrabajoReferenciaFamiliar3��$	��135��$	��Top=�$
��Left=�$��635,130,0,0q�$ Upper=��� $ False��$��TEL. TRABAJO��$	��Top=�$
��Left=�$��640,134,0,0q�
����������� " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$+ %txtTelefonoTrabajoReferenciaFamiliar3��  .a���  � y�$ Visible=��� -�   $��tabFinanciamiento��$��Q��.��� ���	#FFE5E5E5.+��� -�   $ label16��$#��CONDICIONES DE FINANCIAMIENTO��$��Center=�$��	10,10,0,0q�$	��Top=�$��16(�
����  �$
��Bold"��� -�   $ txtArticulo��$	��320��$	��Top=�$
��Left=�$��
240,50,0,0q�$ Upper=��� $ False��$��	ARTÍCULO��$	��Top=�$
��Left=�$��
245,54,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtArticulo��  .a���  � y�$ Visible=��� -�   $ 	txtPrecio��$	��135��$	��Top=�$
��Left=�$��
240,90,0,0q�$ Upper=�$��Right=��� $ False��$��	PRECIO Q.��$	��Top=�$
��Left=�$��
245,94,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ 	txtPrecio��  .a���  � y�$ Visible=��� -�   $ txtPlazo��$	��135��$	��Top=�$
��Left=�$��
430,90,0,0q�$ Upper=�$��Right=��� $ False��$��PLAZO CUOTAS��$	��Top=�$
��Left=�$��
435,94,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtPlazo��  .a���  � y�$ Visible=��� -�   $ txtEnganche��$	��135��$	��Top=�$
��Left=�$��240,130,0,0q�$ Upper=�$��Right=��� $ False��$��ENGANCHE Q.��$	��Top=�$
��Left=�$��245,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtEnganche��  .a���  � y�$ Visible=��� -�   $ txtValorCuota��$	��135��$	��Top=�$
��Left=�$��430,130,0,0q�$ Upper=�$��Right=��� $ False��$��VALOR CUOTA��$	��Top=�$
��Left=�$��435,134,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtValorCuota��  .a���  � y�$ Visible=��� -�   $ txtSaldo��$	��135��$	��Top=�$
��Left=�$��240,170,0,0q�$ Upper=�$��Right=��� $ False��$��SALDO Q.��$	��Top=�$
��Left=�$��245,174,0,0q�
���  ������ " ��.���  � y�$ 	Collapsed=� ��  �� *Text+$ txtSaldo��  .a���  � y�$ Visible=��� -�   $ cmbTipoCredito��$��430,170,0,0q�$	��Top=�$	��135��$
��Left=�.@��� $
	 True��.��TIPO DE CREDITOS�� -�   $��tabComentarios��$��?��.���� -�   $ ScrollComentarios���� -�   $��	tabAcceso��$��@��.��� ���	#FFE5E5E5.+��� -�   $ label14��$��Usuario:��$
��Left=�$��	10,10,0,0q�$	��Top=��� -�   $ 
txtUsuario��$
��Left=�$��	80,10,0,0q�$�� �� �� TextWrapping$
 Wrap=�$	��Top=�$	��120���� -�   $ label15��$��Contraseña:��$
��Left=�$��
230,10,0,0q�$	��Top=�G� -�    G�Name$ txtClave��$
��Left=�$��
320,10,0,0q�$	��Top=�$	��120�� G� PasswordChar$ *���� -�   $
 btnLogin��$��Login��$
��Left=�$��
460,10,0,0q�$	��Top=��� -�   $ lblnombreusuario��$��Nombre Usuario:��$
��Left=�$��
520,10,0,0q�$	��Top=�$��SemiBold"�$y�Hidden=�
��� � ��� -�   $ lblareausuario��$��Area Usuario:��$
��Left=�$��
520,30,0,0q�$	��Top=�$��SemiBold"�$y�Hidden=�
��� � �