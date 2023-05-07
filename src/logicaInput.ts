//Especificamos el tipo y accedemos al elemento
let inputArchivoExcel = document.getElementById('excel-input') as HTMLInputElement;

//Accedemos a la parte en donde mostramos el nombre del archivo
let nombreArchivo = document.getElementById('file-name-id') as HTMLElement;

//Debemos acceder al nombre de los archivos
if(inputArchivoExcel){
  //Accedemos a los archivos y los convertimos en array
  let archivosSeleccionados = inputArchivoExcel.files as any;
  archivosSeleccionados = [...archivosSeleccionados];

  //Guardamos los nombre de los archivos
  let nombreArchivos: string[] = archivosSeleccionados.map((nombreArchivo: File) => {
    return nombreArchivo.name;
  });

  //Mostramos el nombre de los archivos
  nombreArchivo.textContent = "";
  nombreArchivos.forEach(nombre => {
    nombreArchivo.textContent += `${nombre}, `;
  });

  console.log('archivosSeleccionados: ', archivosSeleccionados[0].name);
}
