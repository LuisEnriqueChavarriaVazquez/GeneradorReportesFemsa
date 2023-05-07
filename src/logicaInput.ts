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

//////////////////////////////////


//////////////////////////////////
//ACCEDEMOS NOMBRE DE ARCHIVOS
/////////////////////////////////

//Debemos acceder a los archivos de los archivos
function getFile() {
  if (inputArchivoExcel) {

    inputArchivoExcel.addEventListener('change', () => {

      //Accedemos a los archivos y los convertimos en array
      let archivosSeleccionadosObject: any = inputArchivoExcel.files as any;
      let archivosSeleccionados = [...archivosSeleccionadosObject];

      //Guardamos los nombre de los archivos
      let nombreArchivos: string[] = archivosSeleccionados.map((nombreArchivo: File) => {
        return nombreArchivo.name;
      });

      //Imprimimos el nombre
      imprimirNombreFile(nombreArchivo, nombreArchivos);

      //Imprimimos los archivos en el contenedor con su nombre
      imprimirNombreFileDentroCaja(showFilesSelectedContainer, nombreArchivos);

      //Leemos los archivos
      archivosSeleccionados.forEach((archivo: File) => {
        leerArchivoExcel(archivo);
      });
    })
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
  if(nombreArchivosArr.length >= 1){
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
  }else if(nombreArchivosArr.length == 0){
    showFilesSelectedEmpty.removeAttribute('style');
  }
}

// Creamos una función para leer el archivo Excel
const leerArchivoExcel = (archivo: File) => {
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
    // Hacemos algo con los datos JSON (en este caso, imprimimos en la consola)
    console.log(datosJSON);
  };

  // Leemos el contenido del archivo como una cadena binaria
  lector.readAsBinaryString(archivo);
};

//Simplemente activamos la funcion de getFile al dar click en el boton
buttonSelect.addEventListener('click', () => {
  inputArchivoExcel.value = "";
  getFile();
});

//////////////////////////////////
