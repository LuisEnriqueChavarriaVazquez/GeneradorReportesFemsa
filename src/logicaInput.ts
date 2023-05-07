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
let datosArchivos: any[][];

//////////////////////////////////


//////////////////////////////////
//ACCEDEMOS NOMBRE DE ARCHIVOS
/////////////////////////////////
//Debemos acceder a los archivos de los archivos
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

      //Imprimimos el nombre
      imprimirNombreFile(nombreArchivo, nombreArchivos);

      //Imprimimos los archivos en el contenedor con su nombre
      imprimirNombreFileDentroCaja(showFilesSelectedContainer, nombreArchivos);

      const promesas = archivosSeleccionados.map((archivo: File) => {
        return leerArchivoExcel(archivo);
      });

      Promise.all(promesas).then((datosArchivos) => {

        console.log('datosArchivos: ', datosArchivos);
        // Convertimos el array a una cadena de texto en formato JSON
        const datosArchivosJSON = JSON.stringify(datosArchivos);
        // Guardamos los datos en el localStorage con la clave "datosArchivos"
        localStorage.setItem('datos_archivos_memoria', datosArchivosJSON);

      }).catch((error) => {
        console.error(error);
      });
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

// Creamos una función que recibe un archivo y devuelve una promesa que se resolverá con los datos JSON del archivo
let leerArchivoExcel = (archivo: File): Promise<any> => {
  return new Promise((resolve, reject) => {
    // Creamos una instancia del FileReader
    const lector = new FileReader();

    // Establecemos una función a ejecutar cuando se complete la carga del archivo
    lector.onload = (evento: any) => {
      // Obtenemos los datos del archivo en formato binario
      const datos = evento.target.result;
      // Parseamos los datos binarios en formato Excel y obtenemos un objeto libro
      const libro = XLSX.read(datos, { type: 'binary' });
      // Obtenemos el nombre de la primera hoja del libro
      const nombreHoja = libro.SheetNames[0];
      // Obtenemos la hoja correspondiente al nombre obtenido anteriormente
      const hoja = libro.Sheets[nombreHoja];
      // Convertimos la hoja a un objeto JSON
      const datosJSON = XLSX.utils.sheet_to_json(hoja, { header: 1 });
      // Resolvemos la promesa con los datos JSON
      resolve(datosJSON);
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

    if(datosConvertidos){

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
        fontSize: 12, // tamaño de fuente
        fontFace: 'Calibri', // tipo de fuente
        row: {
          fill: 'F7F7F7', // color de fondo de las filas
          color: '595959', // color de texto de las filas
          bold: true, // texto en negrita
        },
        headerRow: {
          fill: 'D8D8D8', // color de fondo de la primera fila (encabezados)
          color: '595959', // color de texto de la primera fila
          bold: true, // texto en negrita
        },
        borders: {
          left: { pt: '10', color: '595959', style: 'solid' }, // borde izquierdo
          right: { pt: '10', color: '595959', style: 'solid' }, // borde derecho
          top: { pt: '10', color: '595959', style: 'solid' }, // borde superior
          bottom: { pt: '10', color: '595959', style: 'solid' }, // borde inferior
        },
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
