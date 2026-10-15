'use client';

import { useEffect, useState, useMemo } from 'react';
import { supabase } from '@/lib/supabase';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

interface Categoria {
  id: string;
  nombre: string;
  tipo: 'ingreso' | 'gasto';
}

interface Contacto {
  id: string;
  nombre: string;
  cif_nif?: string;
  tipo: 'cliente' | 'proveedor';
  telefono?: string;
  email?: string;
}

interface Movimiento {
  id: string;
  fecha: string;
  tipo: 'ingreso' | 'gasto';
  importe: number;
  concepto: string;
  metodo_pago: 'efectivo' | 'tarjeta' | 'transferencia';
  categoria_id: string;
  contacto_id?: string;
  categorias?: { nombre: string };
  contactos?: { nombre: string };
}

interface Producto {
  id: string;
  nombre: string;
  sku: string;
  stock_actual: number;
  stock_minimo: number;
  precio_compra: number;
  precio_venta: number;
}

// ==========================================
// FUNCIONES DE VALIDACIÓN
// ==========================================

function validarDocumentoIdentidad(documento: string): { valido: boolean; mensaje?: string } {
  const doc = documento.trim().toUpperCase();
  if (!doc) return { valido: true };

  const regexDni = /^(\d{8})([A-Z])$/;
  const regexNie = /^[XYZ]\d{7}[A-Z]$/;
  const regexCif = /^[ABCDEFGHJNPQRSUVW]\d{7}[0-9A-J]$/;
  const letrasDni = 'TRWAGMYFPDXBNJZSQVHLCKE';

  if (regexDni.test(doc)) {
    const match = doc.match(regexDni);
    if (!match) return { valido: false, mensaje: 'DNI no válido.' };
    const numero = parseInt(match[1], 10);
    const letra = match[2];
    const letraCalculada = letrasDni[numero % 23];
    if (letra !== letraCalculada) {
      return { valido: false, mensaje: `Letra de DNI incorrecta. Para ${match[1]} corresponde la letra ${letraCalculada}.` };
    }
    return { valido: true };
  }

  if (regexNie.test(doc)) {
    const niePrefix = doc.charAt(0);
    const nieNumero = doc.substring(1, 8);
    const letra = doc.charAt(8);
    const prefijoNum = niePrefix === 'X' ? '0' : niePrefix === 'Y' ? '1' : '2';
    const numeroCompleto = parseInt(prefijoNum + nieNumero, 10);
    const letraCalculada = letrasDni[numeroCompleto % 23];
    if (letra !== letraCalculada) {
      return { valido: false, mensaje: `Letra de NIE incorrecta. Corresponde ${letraCalculada}.` };
    }
    return { valido: true };
  }

  if (regexCif.test(doc)) {
    return { valido: true };
  }

  return { valido: false, mensaje: 'Formato no válido. Debe ser DNI, NIE o CIF de empresa.' };
}

function validarTelefono(tel: string): { valido: boolean; mensaje?: string } {
  const t = tel.trim();
  if (!t) return { valido: true };

  const soloDigitos = t.replace(/[\s\-\(\)]/g, '');
  const regexInternacional = /^\+?[0-9]{7,15}$/;

  if (!regexInternacional.test(soloDigitos)) {
    return { 
      valido: false, 
      mensaje: 'Introduce un teléfono válido (entre 7 y 15 dígitos con prefijo opcional).' 
    };
  }
  return { valido: true };
}

function validarEmail(correo: string): { valido: boolean; mensaje?: string } {
  const e = correo.trim();
  if (!e) return { valido: true };
  const regex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  if (!regex.test(e)) {
    return { valido: false, mensaje: 'Introduce un correo electrónico válido.' };
  }
  return { valido: true };
}

export default function Home() {
  const [ahora, setAhora] = useState<Date | null>(null);

  useEffect(() => {
    setAhora(new Date());
  }, []);

  const [tab, setTab] = useState<'tesoreria' | 'inventario' | 'contactos'>('tesoreria');
  const [loading, setLoading] = useState(true);

  // Datos principales
  const [categorias, setCategorias] = useState<Categoria[]>([]);
  const [movimientos, setMovimientos] = useState<Movimiento[]>([]);
  const [productos, setProductos] = useState<Producto[]>([]);
  const [contactos, setContactos] = useState<Contacto[]>([]);
  
  // Filtros de Tesorería
  const [filtroTipo, setFiltroTipo] = useState<'todos' | 'ingreso' | 'gasto'>('todos');
  const [periodoFiltro, setPeriodoFiltro] = useState<'mes_actual' | 'mes_anterior' | '1t' | '2t' | '3t' | '4t' | 'todos'>('mes_actual');

  // Formulario Tesorería
  const [tipoMov, setTipoMov] = useState<'ingreso' | 'gasto'>('ingreso');
  const [concepto, setConcepto] = useState('');
  const [importe, setImporte] = useState('');
  const [categoriaId, setCategoriaId] = useState('');
  const [contactoId, setContactoId] = useState('');
  const [metodoPago, setMetodoPago] = useState<'efectivo' | 'tarjeta' | 'transferencia'>('efectivo');
  const [guardandoMov, setGuardandoMov] = useState(false);

  // Vinculación Producto con Venta
  const [productoSeleccionadoId, setProductoSeleccionadoId] = useState('');
  const [cantidadVenta, setCantidadVenta] = useState('1');

  // Formulario Inventario
  const [nombreProd, setNombreProd] = useState('');
  const [sku, setSku] = useState('');
  const [stockActual, setStockActual] = useState('');
  const [stockMinimo, setStockMinimo] = useState('5');
  const [precioCompra, setPrecioCompra] = useState('');
  const [precioVenta, setPrecioVenta] = useState('');
  const [guardandoProd, setGuardandoProd] = useState(false);

  // Formulario Contactos
  const [nombreContacto, setNombreContacto] = useState('');
  const [tipoContacto, setTipoContacto] = useState<'cliente' | 'proveedor'>('cliente');
  const [cifNif, setCifNif] = useState('');
  const [telefono, setTelefono] = useState('');
  const [email, setEmail] = useState('');
  const [erroresContacto, setErroresContacto] = useState<{ nif?: string; tel?: string; email?: string }>({});
  const [guardandoContacto, setGuardandoContacto] = useState(false);

  async function cargarDatos() {
    setLoading(true);
    const [resCat, resMov, resProd, resCont] = await Promise.all([
      supabase.from('categorias').select('*'),
      supabase.from('movimientos').select('*, categorias(nombre), contactos(nombre)').order('fecha', { ascending: false }),
      supabase.from('productos').select('*').order('nombre', { ascending: true }),
      supabase.from('contactos').select('*').order('nombre', { ascending: true })
    ]);

    if (resCat.data) {
      setCategorias(resCat.data);
      if (resCat.data.length > 0 && !categoriaId) {
        const primeraCat = resCat.data.find(c => c.tipo === 'ingreso');
        if (primeraCat) setCategoriaId(primeraCat.id);
      }
    }
    if (resMov.data) setMovimientos(resMov.data as Movimiento[]);
    if (resProd.data) setProductos(resProd.data as Producto[]);
    if (resCont.data) setContactos(resCont.data as Contacto[]);
    setLoading(false);
  }

  useEffect(() => {
    cargarDatos();
  }, []);

  const handleTipoChange = (nuevoTipo: 'ingreso' | 'gasto') => {
    setTipoMov(nuevoTipo);
    setProductoSeleccionadoId('');
    const cat = categorias.find(c => c.tipo === nuevoTipo);
    if (cat) setCategoriaId(cat.id);
  };

  const handleSeleccionarProducto = (prodId: string, cantStr: string = cantidadVenta) => {
    setProductoSeleccionadoId(prodId);
    if (!prodId) return;

    const prod = productos.find(p => p.id === prodId);
    if (prod) {
      const cant = parseInt(cantStr) || 1;
      setConcepto(`Venta: ${prod.nombre} (x${cant})`);
      setImporte((Number(prod.precio_venta) * cant).toFixed(2));
      
      const catVentas = categorias.find(c => c.tipo === 'ingreso' && c.nombre.toLowerCase().includes('venta'));
      if (catVentas) setCategoriaId(catVentas.id);
    }
  };

  const handleCambiarCantidad = (cantStr: string) => {
    setCantidadVenta(cantStr);
    if (productoSeleccionadoId) {
      handleSeleccionarProducto(productoSeleccionadoId, cantStr);
    }
  };

  const handleGuardarMovimiento = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!concepto || !importe || !categoriaId) return;

    let productoAfectado: Producto | undefined;
    const unidadesVenta = parseInt(cantidadVenta) || 1;

    if (tipoMov === 'ingreso' && productoSeleccionadoId) {
      productoAfectado = productos.find(p => p.id === productoSeleccionadoId);
      if (!productoAfectado) return;

      if (productoAfectado.stock_actual < unidadesVenta) {
        alert(`Stock insuficiente. Solo quedan ${productoAfectado.stock_actual} unidades disponibles de "${productoAfectado.nombre}".`);
        return;
      }
    }

    setGuardandoMov(true);

    const { error: errorMov } = await supabase.from('movimientos').insert([
      {
        tipo: tipoMov,
        concepto,
        importe: parseFloat(importe),
        categoria_id: categoriaId,
        contacto_id: contactoId || null,
        metodo_pago: metodoPago,
      }
    ]);

    if (errorMov) {
      alert('Error al registrar movimiento: ' + errorMov.message);
      setGuardandoMov(false);
      return;
    }

    if (productoAfectado) {
      const nuevoStock = productoAfectado.stock_actual - unidadesVenta;
      const { error: errorStock } = await supabase
        .from('productos')
        .update({ stock_actual: nuevoStock })
        .eq('id', productoAfectado.id);

      if (errorStock) {
        alert('Movimiento guardado, pero hubo un error al actualizar el stock: ' + errorStock.message);
      }
    }

    setConcepto('');
    setImporte('');
    setContactoId('');
    setProductoSeleccionadoId('');
    setCantidadVenta('1');
    await cargarDatos();
    setGuardandoMov(false);
  };

  const handleEliminarMovimiento = async (id: string, conceptoTexto: string, importeValor: number) => {
    const confirmar = confirm(`¿Estás seguro de que deseas eliminar el registro "${conceptoTexto}" (${importeValor.toFixed(2)} €)?`);
    if (!confirmar) return;

    const { error } = await supabase.from('movimientos').delete().eq('id', id);
    if (error) {
      alert('No se pudo eliminar el movimiento: ' + error.message);
    } else {
      cargarDatos();
    }
  };

  const handleGuardarProducto = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!nombreProd || !stockActual) return;

    setGuardandoProd(true);
    const { error } = await supabase.from('productos').insert([
      {
        nombre: nombreProd,
        sku: sku || null,
        stock_actual: parseInt(stockActual),
        stock_minimo: parseInt(stockMinimo) || 5,
        precio_compra: parseFloat(precioCompra) || 0,
        precio_venta: parseFloat(precioVenta) || 0,
      }
    ]);

    if (error) {
      alert('Error: ' + error.message);
    } else {
      setNombreProd('');
      setSku('');
      setStockActual('');
      setPrecioCompra('');
      setPrecioVenta('');
      cargarDatos();
    }
    setGuardandoProd(false);
  };

  const handleGuardarContacto = async (e: React.FormEvent) => {
    e.preventDefault();

    const checkDoc = validarDocumentoIdentidad(cifNif);
    const checkTel = validarTelefono(telefono);
    const checkMail = validarEmail(email);

    const nuevosErrores: { nif?: string; tel?: string; email?: string } = {};

    if (!checkDoc.valido) nuevosErrores.nif = checkDoc.mensaje;
    if (!checkTel.valido) nuevosErrores.tel = checkTel.mensaje;
    if (!checkMail.valido) nuevosErrores.email = checkMail.mensaje;

    setErroresContacto(nuevosErrores);

    if (Object.keys(nuevosErrores).length > 0) return;
    if (!nombreContacto) return;

    setGuardandoContacto(true);
    const { error } = await supabase.from('contactos').insert([
      {
        nombre: nombreContacto,
        tipo: tipoContacto,
        cif_nif: cifNif ? cifNif.trim().toUpperCase() : null,
        telefono: telefono ? telefono.trim() : null,
        email: email ? email.trim() : null,
      }
    ]);

    if (error) {
      alert('Error: ' + error.message);
    } else {
      setNombreContacto('');
      setCifNif('');
      setTelefono('');
      setEmail('');
      setErroresContacto({});
      cargarDatos();
    }
    setGuardandoContacto(false);
  };

  const handleEliminarContacto = async (id: string, nombre: string) => {
    const confirmar = confirm(`¿Estás seguro de que deseas eliminar el contacto "${nombre}"?`);
    if (!confirmar) return;

    const { error } = await supabase.from('contactos').delete().eq('id', id);
    if (error) {
      alert('No se pudo eliminar el contacto: ' + error.message);
    } else {
      cargarDatos();
    }
  };

  const actualizarStockRapido = async (id: string, nuevoStock: number) => {
    if (nuevoStock < 0) return;
    const { error } = await supabase.from('productos').update({ stock_actual: nuevoStock }).eq('id', id);
    if (!error) {
      setProductos(productos.map(p => p.id === id ? { ...p, stock_actual: nuevoStock } : p));
    }
  };

  // ==========================================
  // FILTRADO POR FECHAS Y PERIODOS CONTABLES
  // ==========================================
  const movimientosPorPeriodo = useMemo(() => {
    if (periodoFiltro === 'todos') return movimientos;
    if (!ahora) return movimientos;

    const anoActual = ahora.getFullYear();
    const mesActual = ahora.getMonth();

    return movimientos.filter((m) => {
      const fechaMov = new Date(m.fecha);
      const anoMov = fechaMov.getFullYear();
      const mesMov = fechaMov.getMonth();

      if (periodoFiltro === 'mes_actual') {
        return anoMov === anoActual && mesMov === mesActual;
      }

      if (periodoFiltro === 'mes_anterior') {
        const fechaMesAnterior = new Date(anoActual, mesActual - 1, 1);
        return anoMov === fechaMesAnterior.getFullYear() && mesMov === fechaMesAnterior.getMonth();
      }

      if (periodoFiltro === '1t') {
        return anoMov === anoActual && mesMov >= 0 && mesMov <= 2;
      }

      if (periodoFiltro === '2t') {
        return anoMov === anoActual && mesMov >= 3 && mesMov <= 5;
      }

      if (periodoFiltro === '3t') {
        return anoMov === anoActual && mesMov >= 6 && mesMov <= 8;
      }

      if (periodoFiltro === '4t') {
        return anoMov === anoActual && mesMov >= 9 && mesMov <= 11;
      }

      return true;
    });
  }, [movimientos, periodoFiltro, ahora]);

  const movimientosFiltrados = useMemo(() => {
    return movimientosPorPeriodo.filter((m) => {
      if (filtroTipo === 'todos') return true;
      return m.tipo === filtroTipo;
    });
  }, [movimientosPorPeriodo, filtroTipo]);

  const totalIngresos = useMemo(() => {
    return movimientosPorPeriodo
      .filter(m => m.tipo === 'ingreso')
      .reduce((acc, m) => acc + Number(m.importe), 0);
  }, [movimientosPorPeriodo]);

  const totalGastos = useMemo(() => {
    return movimientosPorPeriodo
      .filter(m => m.tipo === 'gasto')
      .reduce((acc, m) => acc + Number(m.importe), 0);
  }, [movimientosPorPeriodo]);

  const beneficioNeto = totalIngresos - totalGastos;
  const productosEnAlerta = productos.filter(p => p.stock_actual <= p.stock_minimo).length;

  const calcSaldoMetodo = (metodo: string) => {
    return movimientosPorPeriodo
      .filter(m => m.metodo_pago === metodo)
      .reduce((acc, m) => acc + (m.tipo === 'ingreso' ? Number(m.importe) : -Number(m.importe)), 0);
  };

  const exportarExcelBonito = async () => {
    if (movimientosPorPeriodo.length === 0) {
      alert('No hay movimientos en el periodo seleccionado para exportar.');
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Cierre de Caja');

    worksheet.columns = [
      { header: 'Fecha', key: 'fecha', width: 14 },
      { header: 'Tipo', key: 'tipo', width: 12 },
      { header: 'Concepto', key: 'concepto', width: 34 },
      { header: 'Contacto (Cliente/Prov)', key: 'contacto', width: 24 },
      { header: 'Categoría', key: 'categoria', width: 26 },
      { header: 'Método de Pago', key: 'metodo', width: 18 },
      { header: 'Importe (€)', key: 'importe', width: 16 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1F2937' },
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.height = 26;

    movimientosPorPeriodo.forEach((m) => {
      const fila = worksheet.addRow({
        fecha: m.fecha,
        tipo: m.tipo.toUpperCase(),
        concepto: m.concepto,
        contacto: m.contactos?.nombre || '—',
        categoria: m.categorias?.nombre || 'General',
        metodo: m.metodo_pago.toUpperCase(),
        importe: m.tipo === 'gasto' ? -Number(m.importe) : Number(m.importe),
      });

      const cellImporte = fila.getCell('importe');
      cellImporte.numFmt = '#,##0.00 €;[Red]-#,##0.00 €';
      cellImporte.font = { bold: true };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `cierre_caja_${periodoFiltro}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-gray-100 font-sans text-gray-800 flex flex-col justify-between">
      
      <div>
        {/* Barra de Navegación */}
        <nav className="bg-white border-b border-gray-200 sticky top-0 z-10 shadow-sm">
          <div className="max-w-6xl mx-auto px-6 h-16 flex items-center justify-between">
            <div className="flex items-center space-x-6">
              
              {/* Branding Impulso Web */}
              <div className="flex items-center gap-3">
                <img
                  src="/logo-impulso.png"
                  alt="Logo Impulso Web"
                  className="w-10 h-10 object-contain rounded-lg bg-white p-1 border border-gray-100 shadow-sm"
                />
                <div>
                  <div className="flex items-center gap-2">
                    <span className="text-lg font-black text-gray-900 tracking-tight">Versión Demo</span>
                    <span className="text-[10px] uppercase font-bold tracking-wider px-2 py-0.5 bg-blue-50 text-blue-700 border border-blue-200 rounded-full">
                      Impulso Web
                    </span>
                  </div>
                  <p className="text-[11px] text-gray-500 hidden sm:block">Control de Caja, Inventario y Tesorería</p>
                </div>
              </div>

              {/* Pestañas de navegación */}
              <div className="flex space-x-2">
                <button
                  onClick={() => setTab('tesoreria')}
                  className={`px-3 py-1.5 rounded-lg text-sm font-medium transition-colors cursor-pointer ${
                    tab === 'tesoreria' ? 'bg-black text-white' : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
                  }`}
                >
                  Tesorería y Caja
                </button>

                <button
                  onClick={() => setTab('inventario')}
                  className={`px-3 py-1.5 rounded-lg text-sm font-medium transition-colors flex items-center gap-1.5 cursor-pointer ${
                    tab === 'inventario' ? 'bg-black text-white' : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
                  }`}
                >
                  Inventario
                  {productosEnAlerta > 0 && (
                    <span className="w-2 h-2 rounded-full bg-red-500 animate-pulse"></span>
                  )}
                </button>

                <button
                  onClick={() => setTab('contactos')}
                  className={`px-3 py-1.5 rounded-lg text-sm font-medium transition-colors cursor-pointer ${
                    tab === 'contactos' ? 'bg-black text-white' : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
                  }`}
                >
                  Contactos ({contactos.length})
                </button>
              </div>
            </div>

            <div className="flex items-center space-x-3">
              {tab === 'tesoreria' && (
                <button
                  onClick={exportarExcelBonito}
                  className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-gray-700 bg-gray-100 hover:bg-gray-200 border border-gray-300 rounded-lg transition-colors cursor-pointer"
                  title="Descargar libro contable filtrado para gestoría"
                >
                  <span>Descargar Excel Gestor</span>
                </button>
              )}
              <div className="flex items-center space-x-2 text-xs">
                <span className="w-2 h-2 rounded-full bg-green-500"></span>
                <span className="text-gray-500 hidden md:inline">Supabase Activo</span>
              </div>
            </div>
          </div>
        </nav>

        {/* Contenido Principal */}
        <main className="max-w-6xl mx-auto p-6 md:p-10 space-y-6">

          {/* ================= PESTAÑA 1: TESORERÍA ================= */}
          {tab === 'tesoreria' && (
            <div className="space-y-6">

              {/* Selector de Periodo */}
              <div className="bg-white p-3 rounded-xl border border-gray-200 shadow-sm flex flex-wrap items-center justify-between gap-3">
                <span className="text-xs font-bold text-gray-500 uppercase tracking-wider">Periodo Contable:</span>
                <div className="flex flex-wrap gap-1 text-xs">
                  <button
                    onClick={() => setPeriodoFiltro('mes_actual')}
                    className={`px-3 py-1.5 rounded-lg font-medium transition-colors cursor-pointer ${
                      periodoFiltro === 'mes_actual' ? 'bg-black text-white shadow' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                    }`}
                  >
                    Mes Actual
                  </button>
                  <button
                    onClick={() => setPeriodoFiltro('mes_anterior')}
                    className={`px-3 py-1.5 rounded-lg font-medium transition-colors cursor-pointer ${
                      periodoFiltro === 'mes_anterior' ? 'bg-black text-white shadow' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                    }`}
                  >
                    Mes Anterior
                  </button>
                  <button
                    onClick={() => setPeriodoFiltro('1t')}
                    className={`px-2.5 py-1.5 rounded-lg font-medium transition-colors cursor-pointer ${
                      periodoFiltro === '1t' ? 'bg-black text-white shadow' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                    }`}
                  >
                    1T
                  </button>
                  <button
                    onClick={() => setPeriodoFiltro('2t')}
                    className={`px-2.5 py-1.5 rounded-lg font-medium transition-colors cursor-pointer ${
                      periodoFiltro === '2t' ? 'bg-black text-white shadow' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                    }`}
                  >
                    2T
                  </button>
                  <button
                    onClick={() => setPeriodoFiltro('3t')}
                    className={`px-2.5 py-1.5 rounded-lg font-medium transition-colors cursor-pointer ${
                      periodoFiltro === '3t' ? 'bg-black text-white shadow' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                    }`}
                  >
                    3T
                  </button>
                  <button
                    onClick={() => setPeriodoFiltro('4t')}
                    className={`px-2.5 py-1.5 rounded-lg font-medium transition-colors cursor-pointer ${
                      periodoFiltro === '4t' ? 'bg-black text-white shadow' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                    }`}
                  >
                    4T
                  </button>
                  <button
                    onClick={() => setPeriodoFiltro('todos')}
                    className={`px-3 py-1.5 rounded-lg font-medium transition-colors cursor-pointer ${
                      periodoFiltro === 'todos' ? 'bg-black text-white shadow' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                    }`}
                  >
                    Todo
                  </button>
                </div>
              </div>
              
              {/* Métricas Principales */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div className="bg-white p-5 rounded-xl border border-gray-200 shadow-sm">
                  <p className="text-xs uppercase font-bold text-gray-400">Total Ingresos</p>
                  <p className="text-2xl font-extrabold text-green-600 mt-1">+{totalIngresos.toFixed(2)} €</p>
                </div>
                <div className="bg-white p-5 rounded-xl border border-gray-200 shadow-sm">
                  <p className="text-xs uppercase font-bold text-gray-400">Total Gastos</p>
                  <p className="text-2xl font-extrabold text-red-600 mt-1">-{totalGastos.toFixed(2)} €</p>
                </div>
                <div className="bg-white p-5 rounded-xl border border-gray-200 shadow-sm">
                  <p className="text-xs uppercase font-bold text-gray-400">Beneficio Neto</p>
                  <p className={`text-2xl font-extrabold mt-1 ${beneficioNeto >= 0 ? 'text-gray-900' : 'text-red-500'}`}>
                    {beneficioNeto.toFixed(2)} €
                  </p>
                </div>
              </div>

              {/* Arqueo de Caja */}
              <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-sm flex flex-wrap items-center justify-between gap-4 text-xs">
                <span className="font-bold text-gray-500 uppercase tracking-wider">Desglose de Caja:</span>
                <div className="flex gap-6">
                  <span>Caja Efectivo: <strong className="text-gray-900 ml-1">{calcSaldoMetodo('efectivo').toFixed(2)} €</strong></span>
                  <span>Datáfono / Tarjeta: <strong className="text-gray-900 ml-1">{calcSaldoMetodo('tarjeta').toFixed(2)} €</strong></span>
                  <span>Banco / Transferencia: <strong className="text-gray-900 ml-1">{calcSaldoMetodo('transferencia').toFixed(2)} €</strong></span>
                </div>
              </div>

              {/* Formulario Movimiento */}
              <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
                <h2 className="text-base font-bold text-gray-900 mb-4">Registrar Movimiento de Caja</h2>
                
                <form onSubmit={handleGuardarMovimiento} className="space-y-4">
                  
                  {tipoMov === 'ingreso' && productos.length > 0 && (
                    <div className="p-3 bg-blue-50 border border-blue-200 rounded-lg flex flex-wrap items-center gap-3 text-xs">
                      <span className="font-bold text-blue-900">¿Es una venta de inventario?</span>
                      
                      <select
                        value={productoSeleccionadoId}
                        onChange={(e) => handleSeleccionarProducto(e.target.value)}
                        className="border rounded-md p-1.5 bg-white text-gray-800 focus:outline-none"
                      >
                        <option value="">(Venta directa / Servicio libre)</option>
                        {productos.map(p => (
                          <option key={p.id} value={p.id}>
                            {p.nombre} — {Number(p.precio_venta).toFixed(2)} € (Stock: {p.stock_actual} uds)
                          </option>
                        ))}
                      </select>

                      {productoSeleccionadoId && (
                        <div className="flex items-center gap-2">
                          <span className="text-blue-900">Cantidad:</span>
                          <input
                            type="number"
                            min="1"
                            value={cantidadVenta}
                            onChange={(e) => handleCambiarCantidad(e.target.value)}
                            className="w-16 border rounded-md p-1.5 bg-white text-center text-gray-800"
                          />
                          <span className="text-[11px] text-blue-700 italic">
                            (Se descontarán automáticamente del almacén)
                          </span>
                        </div>
                      )}
                    </div>
                  )}

                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-6 gap-3">
                    <div className="md:col-span-1">
                      <label className="block text-xs font-semibold text-gray-600 mb-1">Tipo</label>
                      <select
                        value={tipoMov}
                        onChange={(e) => handleTipoChange(e.target.value as 'ingreso' | 'gasto')}
                        className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                      >
                        <option value="ingreso">Ingreso (+)</option>
                        <option value="gasto">Gasto (-)</option>
                      </select>
                    </div>

                    <div className="md:col-span-2">
                      <label className="block text-xs font-semibold text-gray-600 mb-1">Concepto</label>
                      <input
                        type="text"
                        placeholder="Ej. Revisión, Albarán piezas..."
                        value={concepto}
                        onChange={(e) => setConcepto(e.target.value)}
                        required
                        className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                      />
                    </div>

                    <div className="md:col-span-1">
                      <label className="block text-xs font-semibold text-gray-600 mb-1">Importe (€)</label>
                      <input
                        type="number"
                        step="0.01"
                        placeholder="0.00"
                        value={importe}
                        onChange={(e) => setImporte(e.target.value)}
                        required
                        className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                      />
                    </div>

                    <div className="md:col-span-1">
                      <label className="block text-xs font-semibold text-gray-600 mb-1">Categoría</label>
                      <select
                        value={categoriaId}
                        onChange={(e) => setCategoriaId(e.target.value)}
                        className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                      >
                        {categorias
                          .filter(c => c.tipo === tipoMov)
                          .map(c => (
                            <option key={c.id} value={c.id}>{c.nombre}</option>
                          ))}
                      </select>
                    </div>

                    <div className="md:col-span-1">
                      <label className="block text-xs font-semibold text-gray-600 mb-1">Contacto Asignado</label>
                      <select
                        value={contactoId}
                        onChange={(e) => setContactoId(e.target.value)}
                        className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                      >
                        <option value="">(Sin asignar / General)</option>
                        {contactos
                          .filter(c => tipoMov === 'ingreso' ? c.tipo === 'cliente' : c.tipo === 'proveedor')
                          .map(c => (
                            <option key={c.id} value={c.id}>{c.nombre}</option>
                          ))}
                      </select>
                    </div>

                    <div className="md:col-span-1">
                      <label className="block text-xs font-semibold text-gray-600 mb-1">Método</label>
                      <select
                        value={metodoPago}
                        onChange={(e) => setMetodoPago(e.target.value as 'efectivo' | 'tarjeta' | 'transferencia')}
                        className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                      >
                        <option value="efectivo">Efectivo</option>
                        <option value="tarjeta">Tarjeta</option>
                        <option value="transferencia">Transferencia</option>
                      </select>
                    </div>

                    <div className="sm:col-span-2 md:col-span-5 flex justify-end items-end mt-2">
                      <button
                        type="submit"
                        disabled={guardandoMov}
                        className="px-5 py-2.5 bg-black hover:bg-gray-800 text-white font-medium text-sm rounded-lg transition-colors disabled:opacity-50 cursor-pointer"
                      >
                        {guardandoMov ? 'Guardando...' : 'Añadir Registro'}
                      </button>
                    </div>
                  </div>
                </form>
              </div>

              {/* Historial con Botón Eliminar */}
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
                <div className="px-6 py-4 border-b border-gray-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h2 className="text-base font-bold text-gray-900">Historial de Caja</h2>
                    <span className="text-xs text-gray-500">
                      Mostrando movimientos del periodo: <strong className="uppercase text-gray-700">{periodoFiltro.replace('_', ' ')}</strong>
                    </span>
                  </div>
                  
                  <div className="flex gap-1.5 text-xs bg-gray-100 p-1 rounded-lg">
                    <button
                      onClick={() => setFiltroTipo('todos')}
                      className={`px-3 py-1 rounded-md font-medium transition-colors cursor-pointer ${filtroTipo === 'todos' ? 'bg-white shadow-sm text-gray-900 font-bold' : 'text-gray-600'}`}
                    >
                      Todos ({movimientosPorPeriodo.length})
                    </button>
                    <button
                      onClick={() => setFiltroTipo('ingreso')}
                      className={`px-3 py-1 rounded-md font-medium transition-colors cursor-pointer ${filtroTipo === 'ingreso' ? 'bg-white shadow-sm text-green-700 font-bold' : 'text-gray-600'}`}
                    >
                      Ingresos
                    </button>
                    <button
                      onClick={() => setFiltroTipo('gasto')}
                      className={`px-3 py-1 rounded-md font-medium transition-colors cursor-pointer ${filtroTipo === 'gasto' ? 'bg-white shadow-sm text-red-700 font-bold' : 'text-gray-600'}`}
                    >
                      Gastos
                    </button>
                  </div>
                </div>

                {loading ? (
                  <p className="p-6 text-sm text-gray-500">Cargando...</p>
                ) : movimientosFiltrados.length === 0 ? (
                  <p className="p-6 text-sm text-gray-500">No hay movimientos registrados para este periodo o filtro.</p>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-gray-50 text-gray-500 uppercase text-xs">
                        <tr>
                          <th className="px-6 py-3">Fecha</th>
                          <th className="px-6 py-3">Concepto</th>
                          <th className="px-6 py-3">Contacto</th>
                          <th className="px-6 py-3">Categoría</th>
                          <th className="px-6 py-3">Método</th>
                          <th className="px-6 py-3 text-right">Importe</th>
                          <th className="px-6 py-3 text-right">Acción</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100">
                        {movimientosFiltrados.map((m) => (
                          <tr key={m.id} className="hover:bg-gray-50">
                            <td className="px-6 py-3 text-gray-500">{m.fecha}</td>
                            <td className="px-6 py-3 font-medium text-gray-900">{m.concepto}</td>
                            <td className="px-6 py-3 text-gray-600">
                              {m.contactos?.nombre ? (
                                <span className="font-semibold text-gray-800">{m.contactos.nombre}</span>
                              ) : (
                                <span className="text-gray-400 text-xs">—</span>
                              )}
                            </td>
                            <td className="px-6 py-3">
                              <span className="inline-block px-2 py-0.5 rounded text-xs bg-gray-100 text-gray-700">
                                {m.categorias?.nombre || 'General'}
                              </span>
                            </td>
                            <td className="px-6 py-3 text-gray-500 capitalize">{m.metodo_pago}</td>
                            <td className={`px-6 py-3 text-right font-bold ${m.tipo === 'ingreso' ? 'text-green-600' : 'text-red-600'}`}>
                              {m.tipo === 'ingreso' ? '+' : '-'}{Number(m.importe).toFixed(2)} €
                            </td>
                            <td className="px-6 py-3 text-right">
                              <button
                                onClick={() => handleEliminarMovimiento(m.id, m.concepto, Number(m.importe))}
                                className="px-2.5 py-1 text-xs font-semibold text-red-600 hover:text-white hover:bg-red-600 border border-red-200 rounded transition-colors cursor-pointer"
                                title="Eliminar este movimiento"
                              >
                                Eliminar
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

            </div>
          )}

          {/* ================= PESTAÑA 2: INVENTARIO ================= */}
          {tab === 'inventario' && (
            <div className="space-y-6">
              {productosEnAlerta > 0 && (
                <div className="p-4 bg-red-50 border border-red-200 rounded-xl flex items-center justify-between text-red-800 text-sm">
                  <span>Hay <strong>{productosEnAlerta}</strong> artículo(s) por debajo del nivel de stock mínimo.</span>
                </div>
              )}

              <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
                <h2 className="text-base font-bold text-gray-900 mb-4">Añadir Artículo al Catálogo</h2>
                
                <form onSubmit={handleGuardarProducto} className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-3">
                  <div className="md:col-span-2">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Nombre del Producto</label>
                    <input
                      type="text"
                      placeholder="Ej. Aceite Sintético 5W30"
                      value={nombreProd}
                      onChange={(e) => setNombreProd(e.target.value)}
                      required
                      className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                    />
                  </div>

                  <div className="md:col-span-1">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Referencia / SKU</label>
                    <input
                      type="text"
                      placeholder="REF-001"
                      value={sku}
                      onChange={(e) => setSku(e.target.value)}
                      className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                    />
                  </div>

                  <div className="md:col-span-1">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Stock Inicial</label>
                    <input
                      type="number"
                      placeholder="0"
                      value={stockActual}
                      onChange={(e) => setStockActual(e.target.value)}
                      required
                      className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                    />
                  </div>

                  <div className="md:col-span-1">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Stock Mínimo</label>
                    <input
                      type="number"
                      placeholder="5"
                      value={stockMinimo}
                      onChange={(e) => setStockMinimo(e.target.value)}
                      className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                    />
                  </div>

                  <div className="md:col-span-2">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Precio Compra (€)</label>
                    <input
                      type="number"
                      step="0.01"
                      placeholder="0.00"
                      value={precioCompra}
                      onChange={(e) => setPrecioCompra(e.target.value)}
                      className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                    />
                  </div>

                  <div className="md:col-span-2">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Precio Venta PVP (€)</label>
                    <input
                      type="number"
                      step="0.01"
                      placeholder="0.00"
                      value={precioVenta}
                      onChange={(e) => setPrecioVenta(e.target.value)}
                      className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                    />
                  </div>

                  <div className="md:col-span-1 flex items-end">
                    <button
                      type="submit"
                      disabled={guardandoProd}
                      className="w-full py-2.5 bg-black hover:bg-gray-800 text-white font-medium text-sm rounded-lg transition-colors disabled:opacity-50 cursor-pointer"
                    >
                      {guardandoProd ? 'Guardando...' : 'Crear'}
                    </button>
                  </div>
                </form>
              </div>

              <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
                <div className="px-6 py-4 border-b border-gray-200">
                  <h2 className="text-base font-bold text-gray-900">Control de Stock e Inventario</h2>
                </div>

                {productos.length === 0 ? (
                  <p className="p-6 text-sm text-gray-500">No hay artículos en el catálogo.</p>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-gray-50 text-gray-500 uppercase text-xs">
                        <tr>
                          <th className="px-6 py-3">Artículo</th>
                          <th className="px-6 py-3">SKU</th>
                          <th className="px-6 py-3">P. Coste</th>
                          <th className="px-6 py-3">PVP</th>
                          <th className="px-6 py-3 text-center">Unidades</th>
                          <th className="px-6 py-3 text-right">Ajuste</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100">
                        {productos.map((p) => {
                          const esAlerta = p.stock_actual <= p.stock_minimo;
                          return (
                            <tr key={p.id} className="hover:bg-gray-50">
                              <td className="px-6 py-3 font-medium text-gray-900">{p.nombre}</td>
                              <td className="px-6 py-3 text-gray-500 text-xs font-mono">{p.sku || '—'}</td>
                              <td className="px-6 py-3 text-gray-500">{Number(p.precio_compra).toFixed(2)} €</td>
                              <td className="px-6 py-3 font-semibold text-gray-800">{Number(p.precio_venta).toFixed(2)} €</td>
                              <td className="px-6 py-3 text-center">
                                <span className={`inline-block px-2.5 py-1 rounded-full text-xs font-bold ${
                                  esAlerta ? 'bg-red-100 text-red-700 border border-red-200' : 'bg-green-100 text-green-700'
                                }`}>
                                  {p.stock_actual} uds {esAlerta && '(Bajo stock)'}
                                </span>
                              </td>
                              <td className="px-6 py-3 text-right space-x-1">
                                <button
                                  onClick={() => actualizarStockRapido(p.id, p.stock_actual - 1)}
                                  className="px-2.5 py-1 bg-gray-100 hover:bg-gray-200 rounded text-xs font-bold text-gray-700 cursor-pointer"
                                >
                                  -1
                                </button>
                                <button
                                  onClick={() => actualizarStockRapido(p.id, p.stock_actual + 1)}
                                  className="px-2.5 py-1 bg-gray-100 hover:bg-gray-200 rounded text-xs font-bold text-gray-700 cursor-pointer"
                                >
                                  +1
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ================= PESTAÑA 3: CONTACTOS ================= */}
          {tab === 'contactos' && (
            <div className="space-y-6">
              
              {/* Formulario Añadir Contacto */}
              <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
                <h2 className="text-base font-bold text-gray-900 mb-4">Añadir Nuevo Contacto (Cliente / Proveedor)</h2>
                
                <form onSubmit={handleGuardarContacto} noValidate className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-6 gap-4">
                  <div className="md:col-span-1">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Tipo</label>
                    <select
                      value={tipoContacto}
                      onChange={(e) => setTipoContacto(e.target.value as 'cliente' | 'proveedor')}
                      className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                    >
                      <option value="cliente">Cliente</option>
                      <option value="proveedor">Proveedor</option>
                    </select>
                  </div>

                  <div className="md:col-span-2">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Nombre o Razón Social *</label>
                    <input
                      type="text"
                      placeholder="Ej. Juan Pérez / Recambios Europa SL"
                      value={nombreContacto}
                      onChange={(e) => setNombreContacto(e.target.value)}
                      required
                      className="w-full border rounded-lg p-2 text-sm bg-gray-50 focus:bg-white"
                    />
                  </div>

                  <div className="md:col-span-1">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">CIF / NIF</label>
                    <input
                      type="text"
                      placeholder="12345678Z o B12345678"
                      value={cifNif}
                      onChange={(e) => {
                        setCifNif(e.target.value);
                        if (erroresContacto.nif) setErroresContacto({ ...erroresContacto, nif: undefined });
                      }}
                      className={`w-full border rounded-lg p-2 text-sm uppercase ${
                        erroresContacto.nif ? 'border-red-500 bg-red-50' : 'bg-gray-50 focus:bg-white'
                      }`}
                    />
                    {erroresContacto.nif && (
                      <p className="text-[11px] text-red-600 mt-1 leading-tight">{erroresContacto.nif}</p>
                    )}
                  </div>

                  <div className="md:col-span-1">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Teléfono (Nacional o Int.)</label>
                    <input
                      type="tel"
                      placeholder="+34 612 345 678 o +49 151 2345"
                      value={telefono}
                      onChange={(e) => {
                        setTelefono(e.target.value);
                        if (erroresContacto.tel) setErroresContacto({ ...erroresContacto, tel: undefined });
                      }}
                      className={`w-full border rounded-lg p-2 text-sm ${
                        erroresContacto.tel ? 'border-red-500 bg-red-50' : 'bg-gray-50 focus:bg-white'
                      }`}
                    />
                    {erroresContacto.tel && (
                      <p className="text-[11px] text-red-600 mt-1 leading-tight">{erroresContacto.tel}</p>
                    )}
                  </div>

                  <div className="md:col-span-1">
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Email</label>
                    <input
                      type="email"
                      placeholder="contacto@empresa.com"
                      value={email}
                      onChange={(e) => {
                        setEmail(e.target.value);
                        if (erroresContacto.email) setErroresContacto({ ...erroresContacto, email: undefined });
                      }}
                      className={`w-full border rounded-lg p-2 text-sm ${
                        erroresContacto.email ? 'border-red-500 bg-red-50' : 'bg-gray-50 focus:bg-white'
                      }`}
                    />
                    {erroresContacto.email && (
                      <p className="text-[11px] text-red-600 mt-1 leading-tight">{erroresContacto.email}</p>
                    )}
                  </div>

                  <div className="md:col-span-6 flex justify-end">
                    <button
                      type="submit"
                      disabled={guardandoContacto}
                      className="px-6 py-2.5 bg-black hover:bg-gray-800 text-white font-medium text-sm rounded-lg transition-colors disabled:opacity-50 cursor-pointer"
                    >
                      {guardandoContacto ? 'Guardando...' : 'Crear Contacto'}
                    </button>
                  </div>
                </form>
              </div>

              {/* Directorio de Contactos */}
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
                <div className="px-6 py-4 border-b border-gray-200">
                  <h2 className="text-base font-bold text-gray-900">Directorio de Clientes y Proveedores</h2>
                </div>

                {contactos.length === 0 ? (
                  <p className="p-6 text-sm text-gray-500">No hay contactos registrados todavía.</p>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-gray-50 text-gray-500 uppercase text-xs">
                        <tr>
                          <th className="px-6 py-3">Nombre</th>
                          <th className="px-6 py-3">Tipo</th>
                          <th className="px-6 py-3">CIF / NIF</th>
                          <th className="px-6 py-3">Teléfono</th>
                          <th className="px-6 py-3">Email</th>
                          <th className="px-6 py-3 text-right">Volumen Acumulado</th>
                          <th className="px-6 py-3 text-right">Acción</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100">
                        {contactos.map((c) => {
                          const totalMovs = movimientos
                            .filter(m => m.contacto_id === c.id)
                            .reduce((acc, m) => acc + Number(m.importe), 0);

                          return (
                            <tr key={c.id} className="hover:bg-gray-50">
                              <td className="px-6 py-3 font-semibold text-gray-900">{c.nombre}</td>
                              <td className="px-6 py-3">
                                <span className={`inline-block px-2 py-0.5 rounded text-xs font-bold uppercase ${
                                  c.tipo === 'cliente' ? 'bg-blue-100 text-blue-700' : 'bg-purple-100 text-purple-700'
                                }`}>
                                  {c.tipo}
                                </span>
                              </td>
                              <td className="px-6 py-3 text-gray-500 font-mono text-xs">{c.cif_nif || '—'}</td>
                              <td className="px-6 py-3 text-gray-600">{c.telefono || '—'}</td>
                              <td className="px-6 py-3 text-gray-600">{c.email || '—'}</td>
                              <td className="px-6 py-3 text-right font-bold text-gray-900">
                                {totalMovs.toFixed(2)} €
                              </td>
                              <td className="px-6 py-3 text-right">
                                <button
                                  onClick={() => handleEliminarContacto(c.id, c.nombre)}
                                  className="px-2.5 py-1 text-xs font-semibold text-red-600 hover:text-white hover:bg-red-600 border border-red-200 rounded transition-colors cursor-pointer"
                                  title="Eliminar este contacto"
                                >
                                  Eliminar
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

            </div>
          )}

        </main>
      </div>

      {/* Pie de Página Corporativo */}
      <footer className="mt-12 py-6 border-t border-gray-200 text-center text-xs text-gray-500 flex flex-col items-center justify-center gap-2">
        <img
          src="/logo-impulso.png"
          alt="Impulso Web"
          className="w-7 h-7 object-contain rounded bg-white p-0.5 border border-gray-100"
        />
        <p>
          Desarrollado y mantenido por <strong className="text-gray-800 font-semibold">Impulso Web</strong> — Soluciones Digitales a Medida
        </p>
      </footer>

    </div>
  );
}