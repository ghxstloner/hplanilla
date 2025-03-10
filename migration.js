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
  const password = 'root';
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
  
  // El formato en Excel es DD/MM/YYYY
  const partes = fechaExcel.split('/');
  if (partes.length !== 3) return null;
  
  // Convertir a formato MySQL YYYY-MM-DD
  return `${partes[2]}-${partes[1]}-${partes[0]}`;
}

// Función para limpiar y convertir el formato de salario
function convertirSalario(salarioExcel) {
  if (!salarioExcel) return 0;
  
  // Eliminar el símbolo B/. y convertir comas a puntos
  return parseFloat(salarioExcel.replace('B/.', '').replace(',', '.').trim());
}

// Función para obtener el código de categoría basado en el tipo de empleado
function obtenerCodCat(tipoEmpleado) {
  if (!tipoEmpleado) return null;
  
  const tipo = tipoEmpleado.toString().toUpperCase().trim();
  
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
      [descripcionCargo.trim()]
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
    const unidadLimpia = unidadAdministrativa.trim();
    
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
  
  const estado = estadoCivilExcel.toString().trim().toLowerCase();
  
  if (estado.includes('soltero')) return 'Soltero/a';
  if (estado.includes('casado')) return 'Casado/a';
  if (estado.includes('divorciado')) return 'Divorciado/a';
  if (estado.includes('unido') || estado.includes('viudo')) return 'Unido';
  
  return estadoCivilExcel;
}

// Función para verificar si un registro debería procesarse (es un egresado)
function esEgresado(fila) {
  if (!fila || !fila['Estado']) return false;
  
  // Convertir a string en caso de que sea otro tipo de dato
  const estado = String(fila['Estado']);
  
  // Verificar si contiene "R" o "Retirado" en cualquier formato
  return estado.includes('R') && estado.includes('Retirado');
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
      cellDates: true,
      cellNF: false,
      cellText: false
    });
    
    const primerHoja = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[primerHoja];
    
    console.log("Leyendo Excel con cabecera en línea 7...");
    
    // Convertir a JSON, especificando que la cabecera está en la línea 7 (índice 6 base 0)
    const data = XLSX.utils.sheet_to_json(worksheet, {
      raw: false,
      defval: '',
      header: 'A',
      range: 6  // Empezar desde la línea 7 (índice 6 base 0)
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
    
    // Mostrar algunos ejemplos de datos
    console.log(`\nEjemplos de datos (primeras 3 filas):`);
    for (let i = 0; i < Math.min(3, dataSinCabeceras.length); i++) {
      console.log(`Fila ${i+1}:`);
      console.log(`  Estado: '${dataSinCabeceras[i][colEstado]}'`);
      console.log(`  Empleado: '${dataSinCabeceras[i][colEmpleado]}'`);
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
        
        // Mostrar que se está procesando este empleado
        console.log(`Procesando egresado: ${ficha} - Estado: '${estadoValor}'`);
        
        // Preparar los datos para la inserción/actualización
        const datosEmpleado = {
          ficha: ficha,
          estado: 'Egresado', // Valor en nompersonal para egresados
          apenom: fila[colNombre] || '',
          sexo: colSexo ? (String(fila[colSexo] || '').includes('Masculino') ? 'M' : 'F') : '',
          estado_civil: colEstadoCivil ? convertirEstadoCivil(fila[colEstadoCivil]) : '',
          direccion: colDireccion ? (fila[colDireccion] || '') : '',
          fecnac: colFechaNac ? convertirFecha(fila[colFechaNac]) : null,
          fecing: colFechaIng ? convertirFecha(fila[colFechaIng]) : null,
          fecharetiro: colFechaRetiro ? convertirFecha(fila[colFechaRetiro]) : null,
          suesal: colSalario ? convertirSalario(fila[colSalario]) : 0,
          sueldopro: colSalario ? convertirSalario(fila[colSalario]) : 0,
          hora_base: colRata ? convertirSalario(fila[colRata]) : 0,
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