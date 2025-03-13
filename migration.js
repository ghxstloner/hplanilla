// Script para actualizar o insertar información de egresados desde Excel a MySQL
// Requisitos: npm install mysql2 xlsx dotenv prompt-sync fs-extra

const mysql = require('mysql2/promise');
const XLSX = require('xlsx');
const fs = require('fs-extra');
const prompt = require('prompt-sync')({ sigint: true });
const path = require('path');

// Función para solicitar las credenciales de la base de datos
function obtenerCredenciales() {
  console.log('=== ACTUALIZACIÓN DE EGRESADOS EN NOMPERSONAL ===');
  
  const usuario = 'root';
  const password = prompt('Ingrese la contraseña de la base de datos: ', { echo: '*' });
  let baseDatos = 'hotelnhp_planilla';
  
  console.log(`\nConectando a la base de datos ${baseDatos} con usuario ${usuario}...`);
  
  return {
    host: 'localhost',
    user: usuario,
    password: password,
    database: baseDatos
  };
}

// Función para convertir el formato de fecha de Excel a MySQL
function convertirFecha(fechaExcel) {
  if (!fechaExcel) return null;
  
  // Convertir a string en caso de que sea un objeto Date o número
  let fechaStr = String(fechaExcel).trim();
  if (!fechaStr || fechaStr === 'undefined' || fechaStr === 'null' || fechaStr.toLowerCase() === '(nulo)') return null;
  
  try {
    // Caso 1: Si ya está en formato MySQL YYYY-MM-DD, devolverlo tal cual
    if (/^\d{4}-\d{2}-\d{2}$/.test(fechaStr)) {
      return fechaStr;
    }
    
    // Caso 2: Formato principal DD/MM/YYYY (estándar español)
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(fechaStr)) {
      const partes = fechaStr.split('/');
      // Asegurarnos de que es una fecha válida
      const dia = parseInt(partes[0], 10);
      const mes = parseInt(partes[1], 10);
      const año = parseInt(partes[2], 10);
      
      if (dia > 0 && dia <= 31 && mes > 0 && mes <= 12 && año > 1900 && año < 2100) {
        return `${partes[2]}-${partes[1]}-${partes[0]}`;
      }
    }
    
    // Caso 3: Formato con guiones DD-MM-YYYY
    if (/^\d{2}-\d{2}-\d{4}$/.test(fechaStr)) {
      const partes = fechaStr.split('-');
      return `${partes[2]}-${partes[1]}-${partes[0]}`;
    }
    
    // Caso 4: Último intento con Date() para otros formatos
    const fecha = new Date(fechaStr);
    if (!isNaN(fecha.getTime())) {
      const year = fecha.getFullYear();
      const month = String(fecha.getMonth() + 1).padStart(2, '0');
      const day = String(fecha.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    
    // Si todas las conversiones fallan, registrar el error y devolver null
    console.error(`No se pudo convertir la fecha: '${fechaExcel}'`);
    return null;
  } catch (error) {
    console.error(`Error al convertir fecha '${fechaExcel}': ${error.message}`);
    return null;
  }
}

// Función mejorada para limpiar y convertir el formato de salario
function convertirSalario(salarioExcel) {
  if (!salarioExcel) return 0;
  
  try {
    // Convertir a string en caso de que sea otro tipo de dato
    let salarioStr = String(salarioExcel).trim();
    console.log(`Procesando salario: '${salarioStr}'`);
    
    // Eliminar el símbolo B/.
    salarioStr = salarioStr.replace('B/.', '');
    
    // Detectar el formato del número
    // Formato español: 1.234,56 (punto como separador de miles, coma como decimal)
    // Formato americano: 1,234.56 (coma como separador de miles, punto como decimal)
    
    // Verificar si tiene formato español (tiene punto antes de coma)
    const formatoEspanol = salarioStr.includes('.') && salarioStr.includes(',') && 
                           salarioStr.indexOf('.') < salarioStr.indexOf(',');
    
    // Verificar si tiene formato americano (tiene coma antes de punto)
    const formatoAmericano = salarioStr.includes('.') && salarioStr.includes(',') && 
                             salarioStr.indexOf(',') < salarioStr.indexOf('.');
    
    if (formatoEspanol) {
      // Es un número con formato español B/.X.XXX,XX
      // Eliminar todos los puntos (separadores de miles)
      salarioStr = salarioStr.replace(/\./g, '');
      // Reemplazar la coma por punto para obtener un número decimal válido
      salarioStr = salarioStr.replace(',', '.');
      console.log(`Formato español detectado, convertido a: ${salarioStr}`);
    } 
    else if (formatoAmericano) {
      // Es un número con formato americano B/.X,XXX.XX
      // Eliminar todas las comas (separadores de miles)
      salarioStr = salarioStr.replace(/,/g, '');
      // El punto decimal ya está en formato correcto
      console.log(`Formato americano detectado, convertido a: ${salarioStr}`);
    }
    else if (salarioStr.includes(',') && !salarioStr.includes('.')) {
      // Solo tiene coma como separador decimal (formato europeo sin miles)
      salarioStr = salarioStr.replace(',', '.');
      console.log(`Formato decimal con coma detectado, convertido a: ${salarioStr}`);
    }
    // Si solo tiene punto como separador decimal, ya está en formato correcto
    
    // Convertir a número y verificar si es válido
    const salarioNum = parseFloat(salarioStr);
    console.log(`Salario convertido final: ${salarioNum}`);
    
    if (isNaN(salarioNum)) {
      console.error(`Error al convertir salario: '${salarioExcel}' no es un valor numérico válido`);
      return 0;
    }
    
    return salarioNum;
  } catch (error) {
    console.error(`Error al procesar salario '${salarioExcel}': ${error.message}`);
    return 0;
  }
}

// Función para convertir la rata por hora
function convertirRata(rataExcel) {
  // Usar la misma lógica que convertirSalario pero sin multiplicar por 100
  return convertirSalario(rataExcel);
}

// Función para procesar nombres completos
function procesarNombre(nombreCompleto) {
  if (!nombreCompleto) {
    return {
      nombres: '',
      nombres2: '',
      apellidos: '',
      apellido_materno: '',
      apenom: ''
    };
  }
  
  try {
    // Normalizar el nombre (eliminar espacios extras, convertir a mayúsculas)
    let nombre = nombreCompleto.trim().toUpperCase();
    
    // Inicializar variables
    let nombres = '';
    let nombres2 = '';
    let apellidos = '';
    let apellido_materno = '';
    
    // Verificar si el nombre tiene formato "APELLIDOS, NOMBRES"
    if (nombre.includes(',')) {
      const partes = nombre.split(',').map(parte => parte.trim());
      
      // La parte antes de la coma son los apellidos
      const partesApellidos = partes[0].split(' ');
      if (partesApellidos.length >= 2) {
        apellidos = partesApellidos[0];
        apellido_materno = partesApellidos.slice(1).join(' ');
      } else {
        apellidos = partes[0];
      }
      
      // La parte después de la coma son los nombres
      const partesNombres = partes[1].split(' ');
      if (partesNombres.length >= 2) {
        nombres = partesNombres[0];
        nombres2 = partesNombres.slice(1).join(' ');
      } else {
        nombres = partes[1];
      }
    } else {
      // Formato "NOMBRES APELLIDOS"
      const palabras = nombre.split(' ').filter(p => p.trim() !== '');
      
      if (palabras.length === 2) {
        // Un nombre, un apellido
        nombres = palabras[0];
        apellidos = palabras[1];
      } else if (palabras.length === 3) {
        // Un nombre, dos apellidos o dos nombres, un apellido
        // Asumimos el caso más común: un nombre, dos apellidos
        nombres = palabras[0];
        apellidos = palabras[1];
        apellido_materno = palabras[2];
      } else if (palabras.length === 4) {
        // Dos nombres, dos apellidos
        nombres = palabras[0];
        nombres2 = palabras[1];
        apellidos = palabras[2];
        apellido_materno = palabras[3];
      } else if (palabras.length > 4) {
        // Casos complejos
        nombres = palabras[0];
        nombres2 = palabras[1];
        apellidos = palabras[2];
        apellido_materno = palabras.slice(3).join(' ');
      } else {
        // Un solo nombre o apellido
        nombres = nombre;
      }
    }
    
    // Formar el apenom en el formato requerido
    let apenom = nombre;
    
    // Retornar el resultado
    return {
      nombres,
      nombres2,
      apellidos,
      apellido_materno,
      apenom
    };
  } catch (error) {
    console.error(`Error al procesar nombre '${nombreCompleto}': ${error.message}`);
    return {
      nombres: '',
      nombres2: '',
      apellidos: '',
      apellido_materno: '',
      apenom: nombreCompleto || ''
    };
  }
}

// Función para obtener el código de categoría basado en el tipo de empleado
function obtenerCodCat(tipoEmpleado) {
  if (!tipoEmpleado) return null;
  
  const tipo = String(tipoEmpleado).toUpperCase().trim();
  
  if (tipo.includes('SINDICATO')) return 1;
  if (tipo.includes('GERENTE') || tipo.includes('ADMINISTRATIVO')) return 2;
  
  return 2; // Por defecto, administrativo
}

// Función para buscar el código de cargo
async function buscarCodCargo(connection, descripcionCargo) {
  if (!descripcionCargo) return null;
  
  try {
    const [rows] = await connection.execute(
      'SELECT cod_car FROM nomcargos WHERE des_car = ?',
      [String(descripcionCargo).trim()]
    );
    
    if (rows.length > 0) {
      return rows[0].cod_car;
    }
  } catch (error) {
    console.error(`Error al buscar código de cargo: ${error.message}`);
  }
  
  return null;
}

// Función para buscar el código de nivel administrativo
async function buscarCodNivel1(connection, unidadAdministrativa) {
  if (!unidadAdministrativa) return null;
  
  try {
    // Eliminar espacios extras y buscar
    const unidadLimpia = String(unidadAdministrativa).trim();
    
    const [rows] = await connection.execute(
      'SELECT codorg FROM nomnivel1 WHERE TRIM(descrip) = ?',
      [unidadLimpia]
    );
    
    if (rows.length > 0) {
      return rows[0].codorg;
    }
  } catch (error) {
    console.error(`Error al buscar código de nivel: ${error.message}`);
  }
  
  return null;
}

// Función para convertir el estado civil al formato requerido
function convertirEstadoCivil(estadoCivilExcel) {
  if (!estadoCivilExcel) return null;
  
  const estado = String(estadoCivilExcel).trim().toLowerCase();
  
  if (estado.includes('soltero')) return 'Soltero/a';
  if (estado.includes('casado')) return 'Casado/a';
  if (estado.includes('divorciado')) return 'Divorciado/a';
  if (estado.includes('unido') || estado.includes('viudo')) return 'Unido';
  
  return estadoCivilExcel;
}

// Función principal
async function main() {
  try {
    // Obtener credenciales de base de datos
    const dbConfig = obtenerCredenciales();
    
    // Crear conexión a la base de datos
    const connection = await mysql.createConnection(dbConfig);
    console.log('Conexión a la base de datos establecida correctamente.');
    
    // Leer el archivo Excel
    const archivoExcel = path.resolve('EGRESADOS.xlsx');
    console.log(`\nLeyendo archivo Excel: ${archivoExcel}`);
    
    if (!fs.existsSync(archivoExcel)) {
      console.error(`ERROR: El archivo ${archivoExcel} no existe.`);
      await connection.end();
      return;
    }
    
    // Leer el archivo Excel con opciones específicas para garantizar mejor compatibilidad
    const workbook = XLSX.readFile(archivoExcel, {
      type: 'binary',
      cellDates: false,  // NO convertir fechas automáticamente
      cellNF: false,
      cellText: true,    // Obtener todas las celdas como texto
    });
    
    const primerHoja = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[primerHoja];
    
    console.log("Leyendo Excel con cabecera en línea 7...");
    
    // Convertir a JSON, especificando que la cabecera está en la línea 7 (índice 6 base 0)
    const data = XLSX.utils.sheet_to_json(worksheet, {
      raw: false,      // Para obtener strings y no valores "crudos"
      defval: '',
      header: 'A',
      range: 6,        // Empezar desde la línea 7 (índice 6 base 0)
      blankrows: false  // Ignorar filas vacías
    });
    
    // Identificar las cabeceras
    const cabeceras = data[0];
    const dataSinCabeceras = data.slice(1);
    
    console.log(`\nCabeceras encontradas:`);
    console.log(cabeceras);
    
    // Mapear índices de columnas
    let colEstado = null;
    let colEmpleado = null;
    let colNombre = null;
    let colSexo = null;
    let colEstadoCivil = null;
    let colDireccion = null;
    let colFechaNac = null;
    let colFechaIng = null;
    let colFechaRetiro = null;
    let colSalario = null;
    let colRata = null;
    let colTipoEmpleado = null;
    let colPlaza = null;
    let colUnidadAdmin = null;
    
    // Encontrar el índice de cada columna
    for (const [key, value] of Object.entries(cabeceras)) {
      if (!value) continue; // Saltar columnas vacías
      
      const valorLower = String(value).toLowerCase().trim();
      
      if (valorLower.includes('estado') && !valorLower.includes('civil')) {
        colEstado = key;
      } else if (valorLower.includes('empleado') && !valorLower.includes('tipo')) {
        colEmpleado = key;
      } else if (valorLower.includes('nombre') || valorLower === 'apenom') {
        colNombre = key;
      } else if (valorLower.includes('sexo')) {
        colSexo = key;
      } else if (valorLower.includes('civil') || valorLower === 'estado_civil') {
        colEstadoCivil = key;
      } else if (valorLower.includes('direcci')) {
        colDireccion = key;
      } else if (valorLower.includes('nac')) {
        colFechaNac = key;
      } else if (valorLower.includes('ingreso')) {
        colFechaIng = key;
      } else if (valorLower.includes('retiro') && valorLower.includes('fecha')) {
        colFechaRetiro = key;
      } else if (valorLower.includes('salario')) {
        colSalario = key;
      } else if (valorLower.includes('rata') || valorLower.includes('hora')) {
        colRata = key;
      } else if (valorLower.includes('tipo')) {
        colTipoEmpleado = key;
      } else if (valorLower.includes('plaza')) {
        colPlaza = key;
      } else if (valorLower.includes('unidad') || valorLower.includes('administrativa')) {
        colUnidadAdmin = key;
      }
    }
    
    // Verificar que se encontraron las columnas necesarias
    if (!colEstado || !colEmpleado) {
      console.error(`ERROR: No se pudieron identificar las columnas necesarias.`);
      console.log(`colEstado: ${colEstado ? 'Encontrado' : 'No encontrado'}`);
      console.log(`colEmpleado: ${colEmpleado ? 'Encontrado' : 'No encontrado'}`);
      await connection.end();
      return;
    }
    
    console.log(`\nColumnas identificadas:`);
    console.log(`Estado: ${colEstado}`);
    console.log(`Empleado: ${colEmpleado}`);
    console.log(`Nombre: ${colNombre}`);
    console.log(`Fecha Retiro: ${colFechaRetiro}`);
    console.log(`Salario: ${colSalario}`);
    console.log(`Rata: ${colRata}`);
    
    // Mostrar algunos ejemplos de datos para verificación
    console.log(`\nEjemplos de datos (primeras 3 filas):`);
    for (let i = 0; i < Math.min(3, dataSinCabeceras.length); i++) {
      console.log(`Fila ${i+1}:`);
      console.log(`  Estado: '${dataSinCabeceras[i][colEstado]}'`);
      console.log(`  Empleado: '${dataSinCabeceras[i][colEmpleado]}'`);
      
      if (colNombre) {
        const nombreOriginal = dataSinCabeceras[i][colNombre];
        const nombreProcesado = procesarNombre(nombreOriginal);
        console.log(`  Nombre Original: '${nombreOriginal}'`);
        console.log(`  Nombre Procesado: `);
        console.log(`    nombres: '${nombreProcesado.nombres}'`);
        console.log(`    nombres2: '${nombreProcesado.nombres2}'`);
        console.log(`    apellidos: '${nombreProcesado.apellidos}'`);
        console.log(`    apellido_materno: '${nombreProcesado.apellido_materno}'`);
        console.log(`    apenom: '${nombreProcesado.apenom}'`);
      }
      
      if (colFechaRetiro) {
        const fechaRetiroOriginal = dataSinCabeceras[i][colFechaRetiro];
        const fechaRetiroConvertida = convertirFecha(fechaRetiroOriginal);
        console.log(`  Fecha Retiro Original: '${fechaRetiroOriginal}'`);
        console.log(`  Fecha Retiro Convertida: '${fechaRetiroConvertida}'`);
      }
      
      if (colSalario) {
        const salarioOriginal = dataSinCabeceras[i][colSalario];
        const salarioConvertido = convertirSalario(salarioOriginal);
        console.log(`  Salario Original: '${salarioOriginal}'`);
        console.log(`  Salario Convertido: ${salarioConvertido}`);
      }
      
      if (colRata) {
        const rataOriginal = dataSinCabeceras[i][colRata];
        const rataConvertida = convertirRata(rataOriginal);
        console.log(`  Rata Original: '${rataOriginal}'`);
        console.log(`  Rata Convertida: ${rataConvertida}`);
      }
    }
    
    // Solicitar confirmación para continuar
    const confirmacion = prompt('\n¿Los datos se ven correctos? ¿Desea continuar? (s/n): ');
    if (confirmacion.toLowerCase() !== 's') {
      console.log('Proceso cancelado por el usuario.');
      await connection.end();
      return;
    }
    
    // Contador para estadísticas
    let actualizados = 0;
    let insertados = 0;
    let errores = 0;
    let ignorados = 0;
    let egresadosEncontrados = 0;
    
    // Procesar cada fila del Excel (saltando la cabecera)
    for (const fila of dataSinCabeceras) {
      // Procesar SOLO si el Estado contiene "R" y "Retirado"
      const estadoValor = String(fila[colEstado] || '');
      
      if (estadoValor.includes('R') && estadoValor.includes('Retirado')) {
        egresadosEncontrados++;
        
        // Obtener el número de empleado
        const ficha = String(fila[colEmpleado] || '').trim();
        if (!ficha) {
          console.error(`ERROR: Fila sin número de empleado (ficha), ignorando...`);
          errores++;
          continue;
        }
        
        // Procesar el nombre completo
        const nombreInfo = procesarNombre(fila[colNombre]);
        
        // Obtener y convertir la fecha de retiro
        let fechaRetiro = null;
        if (colFechaRetiro) {
          const fechaRetiroOriginal = fila[colFechaRetiro];
          if (!fechaRetiroOriginal || String(fechaRetiroOriginal).toLowerCase() === '(nulo)') {
            fechaRetiro = null;
          } else {
            fechaRetiro = convertirFecha(fechaRetiroOriginal);
          }
        }
        
        // Convertir salario y rata por hora
        const salario = convertirSalario(fila[colSalario]);
        const rataPorHora = convertirRata(fila[colRata]);
        
        // Mostrar que se está procesando este empleado
        console.log(`Procesando egresado: ${ficha} - Nombre: '${nombreInfo.apenom}' - Salario: ${salario}`);
        
        // Preparar los datos para la inserción/actualización
        const datosEmpleado = {
          ficha: ficha,
          estado: 'Egresado', // Valor en nompersonal para egresados
          apenom: nombreInfo.apenom,
          nombres: nombreInfo.nombres,
          nombres2: nombreInfo.nombres2,
          apellidos: nombreInfo.apellidos,
          apellido_materno: nombreInfo.apellido_materno,
          sexo: colSexo ? (String(fila[colSexo] || '').includes('Masculino') ? 'Masculino' : 'Femenino') : '',
          estado_civil: colEstadoCivil ? convertirEstadoCivil(fila[colEstadoCivil]) : '',
          direccion: colDireccion ? (fila[colDireccion] || '') : '',
          fecnac: colFechaNac ? convertirFecha(fila[colFechaNac]) : null,
          fecing: colFechaIng ? convertirFecha(fila[colFechaIng]) : null,
          fecharetiro: fechaRetiro,
          suesal: salario,
          sueldopro: salario,
          hora_base: rataPorHora,
          codcat: colTipoEmpleado ? obtenerCodCat(fila[colTipoEmpleado]) : 2
        };
        
        // Buscar el código de cargo y nivel administrativo solo si se encuentran las columnas
        if (colPlaza) {
          datosEmpleado.codcargo = await buscarCodCargo(connection, fila[colPlaza]);
        }
        
        if (colUnidadAdmin) {
          datosEmpleado.codnivel1 = await buscarCodNivel1(connection, fila[colUnidadAdmin]);
        }
        
        try {
          // Verificar si el empleado ya existe en la base de datos
          const [existente] = await connection.execute(
            'SELECT ficha FROM nompersonal WHERE ficha = ?',
            [ficha]
          );
          
          if (existente.length > 0) {
            // Actualizar el registro existente
            const camposActualizar = Object.entries(datosEmpleado)
              .map(([campo, valor]) => `${campo} = ?`)
              .join(', ');
            
            const valoresActualizar = Object.values(datosEmpleado);
            
            await connection.execute(
              `UPDATE nompersonal SET ${camposActualizar} WHERE ficha = ?`,
              [...valoresActualizar, ficha]
            );
            
            actualizados++;
            console.log(`Actualizado: Empleado ${ficha}`);
          } else {
            // Insertar un nuevo registro
            const campos = Object.keys(datosEmpleado).join(', ');
            const placeholders = Object.values(datosEmpleado).map(() => '?').join(', ');
            
            await connection.execute(
              `INSERT INTO nompersonal (${campos}) VALUES (${placeholders})`,
              Object.values(datosEmpleado)
            );
            
            insertados++;
            console.log(`Insertado: Empleado ${ficha}`);
          }
        } catch (error) {
          console.error(`ERROR al procesar empleado ${ficha}: ${error.message}`);
          errores++;
        }
      } else {
        ignorados++;
      }
    }
    
    // Mostrar estadísticas finales
    console.log('\n=== RESUMEN DE LA OPERACIÓN ===');
    console.log(`Total de registros procesados: ${dataSinCabeceras.length}`);
    console.log(`Egresados encontrados: ${egresadosEncontrados}`);
    console.log(`Registros actualizados: ${actualizados}`);
    console.log(`Registros insertados: ${insertados}`);
    console.log(`Registros ignorados (no egresados): ${ignorados}`);
    console.log(`Errores: ${errores}`);
    
    // Cerrar la conexión a la base de datos
    await connection.end();
    console.log('\nProceso completado. Conexión cerrada.');
    
  } catch (error) {
    console.error(`ERROR GENERAL: ${error.message}`);
    console.error(error.stack);
  }
}

// Ejecutar la función principal
main();