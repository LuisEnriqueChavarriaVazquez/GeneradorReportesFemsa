//////////////////////////////////
//ACCESO A LOS ELEMENTOS
/////////////////////////////////

//Especificamos el tipo y accedemos al elemento
let inputArchivoExcel = document.getElementById('excel-input') as HTMLInputElement;

//Accedemos a la parte en donde mostramos el nombre del archivo
let nombreArchivo = document.getElementById('file-name-id') as HTMLElement;

//Accedemos al boton para la seleccion de archivos
let buttonSelect = document.getElementById('buttonSelect') as HTMLElement;

//Accedemos al contenedor para mostrar los archivos elegidos y cuando no tenemos nada
let showFilesSelectedContainer = document.getElementById('showFilesSelectedContainer') as HTMLElement;
let showFilesSelectedEmpty = document.getElementById('showFilesSelectedEmpty') as HTMLElement;

//Archivos de forma global
let archivos: any;
let datosArchivos: any[][]; //Estos serán guardados en la memoria

//////////////////////////////////


//////////////////////////////////
//FUNCIONES DE PANTALLA DE CARGA
/////////////////////////////////
const mostrarCarga = () => {
  const elementoCarga = document.querySelector('#loaderContainer');
  if (elementoCarga) {
    elementoCarga.setAttribute('style', 'display:flex;');
  }
};

const ocultarCarga = () => {
  const elementoCarga = document.querySelector('#loaderContainer');
  if (elementoCarga) {
    elementoCarga.setAttribute('style', 'display:none;');;
  }
};



//////////////////////////////////
//ACCEDEMOS NOMBRE DE ARCHIVOS
/////////////////////////////////
//Debemos acceder a los archivos
function getFile() {
  if (inputArchivoExcel) {
    inputArchivoExcel.addEventListener('change', () => {

      //Accedemos a los archivos y los convertimos en array
      const archivosSeleccionadosObject: any = inputArchivoExcel.files as any;
      const archivosSeleccionados = [...archivosSeleccionadosObject];

      //Guardamos los nombre de los archivos
      const nombreArchivos: string[] = archivosSeleccionados.map((nombreArchivo: File) => {
        return nombreArchivo.name;
      });

      //Imprimimos el nombre abajo del boton
      imprimirNombreFile(nombreArchivo, nombreArchivos);

      //Imprimimos los archivos en el contenedor con su nombre y una tarjeta
      imprimirNombreFileDentroCaja(showFilesSelectedContainer, nombreArchivos);

      /////Obtención de los datos del archivo
      // Creamos un array de promesas para leer los archivos seleccionados.
      const promesas = archivosSeleccionados.map((archivo: File) => {
        return leerArchivoExcel(archivo);
      });

      // Utilizamos Promise.all para esperar a que todas las promesas se resuelvan.
      Promise.all(promesas).then((datosArchivos) => {

        console.log("Todos los datos: ", datosArchivos[0]) // [[hoja1],[hoja2]]

        //Convertimos los datos de las hojas en variables del localstorage
        for (let i = 0; i < datosArchivos[0].length; i++) {
          // Convertimos el array de datos a una cadena de texto en formato JSON.
          const datosArchivosJSON = JSON.stringify(datosArchivos[0][i]);
          // Guardamos los datos en el localStorage con la clave "datosArchivos". Para poder accederlo despues
          localStorage.setItem(`hoja${i + 1}`, datosArchivosJSON);
        }
        //Guardamos la longitud de las hojas
        localStorage.setItem('longitud', datosArchivos[0].length)

      }).catch((error) => {
        // Si alguna de las promesas falla, mostramos el error en la consola.
        console.error(error);
      });

      //Hacemos visible la seccion de configuraciones de la presentación
      mostrarSeccionConfiguraciones();

      //Imprimimos en la pantalla el numero de hojas disponibles en el documento
      imprimirNumeroHojas();
    });
  }
}


//Esta funcion imprime el nombre de los archivos elegidos
function imprimirNombreFile(nombreFileContainer: HTMLElement, nombreArchivosArr: string[]) {
  //Borramos el texto anterior
  nombreFileContainer.textContent = "";

  //En caso de tener mas de un elemento ponemos coma
  if (nombreArchivosArr.length > 1) {
    //Metemos los elementos
    nombreArchivosArr.forEach((nombre: string) => {
      nombreFileContainer.textContent += `${nombre}, `;
    });
  } else if (nombreArchivosArr.length == 1) { //Si tenemos solo uno va sin coma
    nombreFileContainer.textContent += `${nombreArchivosArr[0]}`;
  }
}

//Esta funcion pone los nombre de los archivos cargados dentro de una caja
function imprimirNombreFileDentroCaja(nombreFileContainer: HTMLElement, nombreArchivosArr: string[]) {
  //Borramos el texto anterior
  nombreFileContainer.innerHTML = "";

  //Metemos los elementos
  if (nombreArchivosArr.length >= 1) {
    nombreArchivosArr.forEach((nombre: string) => {
      nombreFileContainer.innerHTML += `
      <div class="showFilesSelected-element border1 shadow1">
      <span class="material-symbols-outlined">
      article
      </span>
      <p>${nombre}</p>
      </div>
      `;
    });
    showFilesSelectedEmpty.setAttribute('style', 'display: none;');
  } else if (nombreArchivosArr.length == 0) {
    showFilesSelectedEmpty.removeAttribute('style');
  }
}

//Imprime en pantalla el numero total de hojas
function imprimirNumeroHojas() {

  setTimeout(() => {

    //Accedemos a los datos del titulo de configuracion y la longitud
    let numeroHojas = document.getElementById('numeroHojas') as HTMLElement;
    let longitud = localStorage.getItem('longitud');

    if(longitud !== null){

      //Insertamos el numero de hojas
      if (numeroHojas) {
        if (longitud == "1") {
          numeroHojas.textContent = longitud + " hoja";
        } else {
          numeroHojas.textContent = longitud + " hojas";
        }
      }

      //Insertamos los inputs de titulo segun el numero de hojas
      let contenedorInputsTitulo = document.getElementById('contenedorInputsTitulo') as HTMLElement;
      for(var i = 0; i < parseInt(longitud); i++){
        contenedorInputsTitulo.innerHTML +=
          `<div class="inputTitleContainer border1 shadow1animated">
            <label for="title_${i+1}" id="title_${i+1}" class="labelTitle">Título de dispositiva ${i+1}</label>
            <input name="title_${i+1}" type="text" class="inputTitle border1" placeholder="Escriba el titulo de la diapositiva.">
          </div>`
        ;
      }

    }

  }, 100);

}

//Funcion para mostrar la sección de configuraciones
function mostrarSeccionConfiguraciones() {
  let seccionConfiguraciones = document.getElementById('seccionConfiguraciones');
  if (seccionConfiguraciones) {
    seccionConfiguraciones.setAttribute('style', 'display: grid; grid-template-columns: 1fr; grid-template-rows: auto; gap: 10px;');
  }
}

//////////////////////////////
//Lectura de los datos
//////////////////////////////
// Creamos una función que recibe un archivo y devuelve una promesa que se resolverá con los datos JSON del archivo
let leerArchivoExcel = (archivo: File): Promise<any> => {
  return new Promise((resolve, reject) => {
    //Mostramos la pantalla de carga
    mostrarCarga();
    // Creamos una instancia del FileReader
    const lector = new FileReader();

    // Establecemos una función a ejecutar cuando se complete la carga del archivo
    lector.onload = (evento: any) => {
      // Obtenemos los datos del archivo en formato binario
      const datos = evento.target.result;
      // Parseamos los datos binarios en formato Excel y obtenemos un objeto libro
      const libro = XLSX.read(datos, { type: 'binary' });
      //Obtenemos la longitud de hojas en el excel
      const numeroHojas = libro.SheetNames.length;

      // Creamos un arreglo para almacenar los datos de todas las hojas
      const datosHoja = [];

      // Recorremos el arreglo de nombres de hojas y convertimos cada hoja a un objeto JSON
      for (let i = 0; i < numeroHojas; i++) {
        const nombreHoja = libro.SheetNames[i];
        const hoja = libro.Sheets[nombreHoja];
        const datosJSON = XLSX.utils.sheet_to_json(hoja, { header: 1 });
        datosHoja.push(datosJSON);
      }

      // Resolvemos la promesa con los datos de todas las hojas y el número de hojas
      ocultarCarga();
      resolve(datosHoja);
    };


    // Establecemos una función a ejecutar en caso de error
    lector.onerror = (evento: any) => {
      // Rechazamos la promesa con el error obtenido
      reject(evento.target.error);
    };

    // Leemos el contenido del archivo como una cadena binaria
    lector.readAsBinaryString(archivo);
  });
};


//Simplemente activamos la funcion de getFile al dar click en el boton
buttonSelect.addEventListener('click', () => {
  inputArchivoExcel.value = "";
  localStorage.clear();
  getFile();
});

//////////////////////////////////


//////////////////////////////////
//Creación del reporte
/////////////////////////////////

//Accedemos al boton para generar los formatos
let createReport = document.getElementById('createReport');
let numeroDiapo = 1;

//Creamos la funcion para agregar el formato
function agregarFormato(slide: any) {
  // Agregamos los encabezados para la diapositiva.
  slide.addShape('rect', {
    x: 0, // posición horizontal en pulgadas
    y: 0, // posición vertical en pulgadas
    w: 10, // ancho en pulgadas
    h: .9, // alto en pulgadas 5.64 es el maximo
    fill: 'D80032', // color de relleno en formato hexadecimal
  });

  slide.addShape('rect', {
    x: 0,
    y: .9,
    w: 10,
    h: .1,
    fill: 'b9002b',
  });

  slide.addShape('rect', {
    x: 0,
    y: 5.61,
    w: 10,
    h: .3,
    fill: 'b9002b',
  });

  // Agrega una imagen desde una URL
  slide.addImage({
    path: '../img/logo.png',
    x: 8.5, // posición horizontal en pulgadas
    y: -.2, // posición vertical en pulgadas
    w: 1.5, // ancho en pulgadas
    h: 1.3 // alto en pulgadas
  });

  //Agregamos la numeracion
  slide.addText(numeroDiapo++, { x: 9.5, y: 5.75, fontSize: 15, color: "ffffff" });
}

//Creamos el reporte
createReport?.addEventListener('click', () => {

  //Son los datos provenientes de los excel que fueron guardados de forma asincrona en memoria
  let datosConvertidos = localStorage.getItem('datos_archivos_memoria');

  if (datosConvertidos !== null) {
    datosConvertidos = JSON.parse(datosConvertidos);
    console.log('datosConvertidos: ', datosConvertidos);
  } else {
    console.log('No se encontró la clave "datos_archivos_memoria" en el almacenamiento local');
  }

  if (datosConvertidos) {

    //Iniciamos la creacion del reporte
    const pptx = new PptxGenJS();

    // Añadimos una nueva diapositiva con un título
    const slide1 = pptx.addSlide({ masterName: 'Primera diapositiva.' });

    agregarFormato(slide1);
    slide1.addText('Título de la diapositiva 1', { x: 0, y: 0.45, fontSize: 22, color: 'ffffff', fontFace: 'Arial', bold: true });

    slide1.addText('Hola!', { x: 1.5, y: 2.0, fontSize: 48 });

    // Agregar la tabla al slide
    slide1.addTable(datosConvertidos[0], {
      x: 0.5, // posición x de la tabla en pulgadas
      y: 1.5, // posición y de la tabla en pulgadas
      w: 8, // ancho de la tabla en pulgadas
      h: 2, // alto de la tabla en pulgadas
      columnWidths: [2, 2, 4], // ancho de cada columna en pulgadas
      autoPageBreaks: true,
      fontFace: "Arial",
      fontSize: 12,
      rowHeight: 0.5,
      align: "center",
      valign: "middle",
      border: {
        pt: 1,
        color: "000000"
      }
    });

    // Añadimos una nueva diapositiva con dos columnas de texto
    const slide2 = pptx.addSlide({ masterName: 'Segunda diapositiva' });

    agregarFormato(slide2);
    slide2.addText('Título de la diapositiva 2', { x: 0, y: 0.45, fontSize: 22, color: 'ffffff', fontFace: 'Arial', bold: true });

    // Añadimos una nueva diapositiva con dos columnas de texto
    const slide3 = pptx.addSlide({ masterName: 'Segunda diapositiva' });

    agregarFormato(slide3);
    slide3.addText('Título de la diapositiva 3', { x: 0, y: 0.45, fontSize: 22, color: 'ffffff', fontFace: 'Arial', bold: true });

    // Generamos la presentación
    pptx.writeFile('ejemplo.pptx');


  }

})
