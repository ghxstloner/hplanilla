const mysql = require('mysql2/promise');
const XLSX = require('xlsx');
const fs = require('fs-extra');
const prompt = require('prompt-sync')({ sigint: true });
const path = require('path');

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

function esValorNulo(valor) {
  if (!valor) return true;
  
  const str = String(valor).trim().toLowerCase();
  return str === '' || str === 'undefined' || str === 'null' || str === '(nulo)';
}

function convertirFecha(fechaExcel) {
  if (esValorNulo(fechaExcel)) return null;
  
  let fechaStr = String(fechaExcel).trim();
  
  try {
    if (/^\d{4}-\d{2}-\d{2}$/.test(fechaStr)) {
      return fechaStr;
    }
    
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(fechaStr)) {
      const partes = fechaStr.split('/');
      const dia = parseInt(partes[0], 10);
      const mes = parseInt(partes[1], 10);
      const año = parseInt(partes[2], 10);
      
      if (dia > 0 && dia <= 31 && mes > 0 && mes <= 12 && año > 1900 && año < 2100) {
        return `${partes[2]}-${partes[1]}-${partes[0]}`;
      }
    }
    
    if (/^\d{2}-\d{2}-\d{4}$/.test(fechaStr)) {
      const partes = fechaStr.split('-');
      return `${partes[2]}-${partes[1]}-${partes[0]}`;
    }
    
    const fecha = new Date(fechaStr);
    if (!isNaN(fecha.getTime())) {
      const year = fecha.getFullYear();
      const month = String(fecha.getMonth() + 1).padStart(2, '0');
      const day = String(fecha.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    
    console.error(`No se pudo convertir la fecha: '${fechaExcel}'`);
    return null;
  } catch (error) {
    console.error(`Error al convertir fecha '${fechaExcel}': ${error.message}`);
    return null;
  }
}

function convertirSalario(salarioExcel) {
  if (esValorNulo(salarioExcel)) return 0;
  
  try {
    let salarioStr = String(salarioExcel).trim();
    console.log(`Procesando salario: '${salarioStr}'`);
    
    salarioStr = salarioStr.replace('B/.', '');
    
    const formatoEspanol = salarioStr.includes('.') && salarioStr.includes(',') && 
                           salarioStr.indexOf('.') < salarioStr.indexOf(',');
    
    const formatoAmericano = salarioStr.includes('.') && salarioStr.includes(',') && 
                             salarioStr.indexOf(',') < salarioStr.indexOf('.');
    
    if (formatoEspanol) {
      salarioStr = salarioStr.replace(/\./g, '');
      salarioStr = salarioStr.replace(',', '.');
      console.log(`Formato español detectado, convertido a: ${salarioStr}`);
    } 
    else if (formatoAmericano) {
      salarioStr = salarioStr.replace(/,/g, '');
      console.log(`Formato americano detectado, convertido a: ${salarioStr}`);
    }
    else if (salarioStr.includes(',') && !salarioStr.includes('.')) {
      salarioStr = salarioStr.replace(',', '.');
      console.log(`Formato decimal con coma detectado, convertido a: ${salarioStr}`);
    }
    
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

function convertirRata(rataExcel) {
  return convertirSalario(rataExcel);
}

function procesarNombre(nombreCompleto) {
  if (esValorNulo(nombreCompleto)) {
    return {
      nombres: null,
      nombres2: null,
      apellidos: null,
      apellido_materno: null,
      apenom: null
    };
  }
  
  try {
    let nombre = nombreCompleto.trim().toUpperCase();
    
    let nombres = null;
    let nombres2 = null;
    let apellidos = null;
    let apellido_materno = null;
    
    if (nombre.includes(',')) {
      const partes = nombre.split(',').map(parte => parte.trim());
      
      const partesApellidos = partes[0].split(' ');
      if (partesApellidos.length >= 2) {
        apellidos = partesApellidos[0];
        apellido_materno = partesApellidos.slice(1).join(' ');
      } else {
        apellidos = partes[0];
      }
      
      const partesNombres = partes[1].split(' ');
      if (partesNombres.length >= 2) {
        nombres = partesNombres[0];
        nombres2 = partesNombres.slice(1).join(' ');
      } else {
        nombres = partes[1];
      }
    } else {
      const palabras = nombre.split(' ').filter(p => p.trim() !== '');
      
      if (palabras.length === 2) {
        nombres = palabras[0];
        apellidos = palabras[1];
      } else if (palabras.length === 3) {
        nombres = palabras[0];
        apellidos = palabras[1];
        apellido_materno = palabras[2];
      } else if (palabras.length === 4) {
        nombres = palabras[0];
        nombres2 = palabras[1];
        apellidos = palabras[2];
        apellido_materno = palabras[3];
      } else if (palabras.length > 4) {
        nombres = palabras[0];
        nombres2 = palabras[1];
        apellidos = palabras[2];
        apellido_materno = palabras.slice(3).join(' ');
      } else {
        nombres = nombre;
      }
    }
    
    // Construir apenom con los componentes disponibles
    let apenom = '';
    if (apellidos) apenom += apellidos;
    if (apellido_materno) apenom += `, ${apellido_materno}`;
    apenom += ', ';
    if (nombres) apenom += nombres;
    if (nombres2) apenom += ` ${nombres2}`;
    
    apenom = apenom.trim();
    // Si termina con una coma, quitarla
    if (apenom.endsWith(',')) apenom = apenom.slice(0, -1);
    
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
      nombres: null,
      nombres2: null,
      apellidos: null,
      apellido_materno: null,
      apenom: nombreCompleto
    };
  }
}

function obtenerCodCat(tipoEmpleado) {
  if (!tipoEmpleado) return null;
  
  const tipo = String(tipoEmpleado).toUpperCase().trim();
  
  if (tipo.includes('SINDICATO')) return 1;
  if (tipo.includes('GERENTE') || tipo.includes('ADMINISTRATIVO')) return 2;
  
  return 2;
}

async function buscarCodCargo(connection, descripcionCargo) {
  if (esValorNulo(descripcionCargo)) return null;
  
  const descripcion = String(descripcionCargo).trim();
  
  try {
    // Intentar encontrar el cargo existente
    const [rows] = await connection.execute(
      'SELECT cod_car FROM nomcargos WHERE des_car = ?',
      [descripcion]
    );
    
    if (rows.length > 0) {
      console.log(`Cargo encontrado: ${descripcion} - Código: ${rows[0].cod_car}`);
      return rows[0].cod_car;
    } else {
      // Si no existe, crear el cargo
      console.log(`Cargo no encontrado: ${descripcion}. Creando nuevo cargo...`);
      
      // Obtener el máximo código de cargo para generar uno nuevo
      const [maxRows] = await connection.execute('SELECT MAX(cod_car) as max_codigo FROM nomcargos');
      const nuevoCodigo = maxRows[0].max_codigo ? maxRows[0].max_codigo + 1 : 1;
      
      // Insertar el nuevo cargo
      await connection.execute(
        'INSERT INTO nomcargos (cod_car, des_car) VALUES (?, ?)',
        [nuevoCodigo, descripcion]
      );
      
      console.log(`Nuevo cargo creado: ${descripcion} - Código: ${nuevoCodigo}`);
      return nuevoCodigo;
    }
  } catch (error) {
    console.error(`Error al buscar/crear código de cargo: ${error.message}`);
    return null;
  }
}

async function buscarCodNivel1(connection, unidadAdministrativa) {
  if (esValorNulo(unidadAdministrativa)) return null;
  
  try {
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

function convertirEstadoCivil(estadoCivilExcel) {
  if (!estadoCivilExcel) return null;
  
  const estado = String(estadoCivilExcel).trim().toLowerCase();
  
  if (estado.includes('soltero')) return 'Soltero/a';
  if (estado.includes('casado')) return 'Casado/a';
  if (estado.includes('divorciado')) return 'Divorciado/a';
  if (estado.includes('unido') || estado.includes('viudo')) return 'Unido';
  
  return estadoCivilExcel;
}

async function main() {
  try {
    const dbConfig = obtenerCredenciales();
    
    const connection = await mysql.createConnection(dbConfig);
    console.log('Conexión a la base de datos establecida correctamente.');
    
    const archivoExcel = path.resolve('EGRESADOS.xlsx');
    console.log(`\nLeyendo archivo Excel: ${archivoExcel}`);
    
    if (!fs.existsSync(archivoExcel)) {
      console.error(`ERROR: El archivo ${archivoExcel} no existe.`);
      await connection.end();
      return;
    }
    
    const workbook = XLSX.readFile(archivoExcel, {
      type: 'binary',
      cellDates: false,
      cellNF: false,
      cellText: true,
    });
    
    const primerHoja = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[primerHoja];
    
    console.log("Leyendo Excel con cabecera en línea 7...");
    
    const data = XLSX.utils.sheet_to_json(worksheet, {
      raw: false,
      defval: '',
      header: 'A',
      range: 6,
      blankrows: false
    });
    
    const cabeceras = data[0];
    const dataSinCabeceras = data.slice(1);
    
    console.log(`\nCabeceras encontradas:`);
    console.log(cabeceras);
    
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
    
    for (const [key, value] of Object.entries(cabeceras)) {
      if (!value) continue;
      
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
    
    const confirmacion = prompt('\n¿Los datos se ven correctos? ¿Desea continuar? (s/n): ');
    if (confirmacion.toLowerCase() !== 's') {
      console.log('Proceso cancelado por el usuario.');
      await connection.end();
      return;
    }
    
    let actualizados = 0;
    let insertados = 0;
    let errores = 0;
    let ignorados = 0;
    let egresadosEncontrados = 0;
    
    for (const fila of dataSinCabeceras) {
      const estadoValor = String(fila[colEstado] || '');
      
      if (estadoValor.includes('R') && estadoValor.includes('Retirado')) {
        egresadosEncontrados++;
        
        const ficha = String(fila[colEmpleado] || '').trim();
        if (!ficha) {
          console.error(`ERROR: Fila sin número de empleado (ficha), ignorando...`);
          errores++;
          continue;
        }
        
        const nombreInfo = procesarNombre(fila[colNombre]);
        
        let fechaRetiro = null;
        if (colFechaRetiro) {
          const fechaRetiroOriginal = fila[colFechaRetiro];
          fechaRetiro = convertirFecha(fechaRetiroOriginal);
        }
        
        const salario = convertirSalario(fila[colSalario]);
        const rataPorHora = convertirRata(fila[colRata]);
        
        console.log(`Procesando egresado: ${ficha} - Nombre: '${nombreInfo.apenom}' - Salario: ${salario}`);
        
        // Procesar los valores para evitar que (nulo) vaya a la base de datos
        const sexoValor = colSexo ? (esValorNulo(fila[colSexo]) ? null : 
                                    (String(fila[colSexo]).includes('Masculino') ? 'Masculino' : 'Femenino')) : null;
        
        const estadoCivilValor = colEstadoCivil ? 
                                (esValorNulo(fila[colEstadoCivil]) ? null : 
                                convertirEstadoCivil(fila[colEstadoCivil])) : null;
        
        const direccionValor = colDireccion ? 
                              (esValorNulo(fila[colDireccion]) ? null : String(fila[colDireccion])) : null;
        
        const datosEmpleado = {
          ficha: ficha,
          estado: 'Egresado',
          apenom: nombreInfo.apenom,
          nombres: nombreInfo.nombres,
          nombres2: nombreInfo.nombres2,
          apellidos: nombreInfo.apellidos,
          apellido_materno: nombreInfo.apellido_materno,
          sexo: sexoValor,
          estado_civil: estadoCivilValor,
          direccion: direccionValor,
          fecnac: colFechaNac ? convertirFecha(fila[colFechaNac]) : null,
          fecing: colFechaIng ? convertirFecha(fila[colFechaIng]) : null,
          fecharetiro: fechaRetiro,
          suesal: salario,
          sueldopro: salario,
          hora_base: rataPorHora,
          codcat: colTipoEmpleado ? obtenerCodCat(fila[colTipoEmpleado]) : 2
        };
        
        if (colPlaza) {
          datosEmpleado.codcargo = await buscarCodCargo(connection, fila[colPlaza]);
        }
        
        if (colUnidadAdmin) {
          datosEmpleado.codnivel1 = await buscarCodNivel1(connection, fila[colUnidadAdmin]);
        }
        
        try {
          const [existente] = await connection.execute(
            'SELECT ficha FROM nompersonal WHERE ficha = ?',
            [ficha]
          );
          
          if (existente.length > 0) {
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
    
    console.log('\n=== RESUMEN DE LA OPERACIÓN ===');
    console.log(`Total de registros procesados: ${dataSinCabeceras.length}`);
    console.log(`Egresados encontrados: ${egresadosEncontrados}`);
    console.log(`Registros actualizados: ${actualizados}`);
    console.log(`Registros insertados: ${insertados}`);
    console.log(`Registros ignorados (no egresados): ${ignorados}`);
    console.log(`Errores: ${errores}`);
    
    await connection.end();
    console.log('\nProceso completado. Conexión cerrada.');
    
  } catch (error) {
    console.error(`ERROR GENERAL: ${error.message}`);
    console.error(error.stack);
  }
}

main();