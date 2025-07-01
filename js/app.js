function getQueryParam(param) {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(param);
}

const id = getQueryParam("id");

fetch("instrumentos.json")
  .then(response => response.json())
  .then(data => {
    const found = data.find(item => item["IDENTIFICACIÓN"] === id);

    if (found) {
      document.getElementById("id").textContent = found["IDENTIFICACIÓN"];
      document.getElementById("equipo").textContent = found["EQUIPO  /  INSTRUMENTO"];
      document.getElementById("fabricante").textContent = found["FABRICANTE"];
      document.getElementById("serie").textContent = found["SERIE "] ?? "";
      document.getElementById("modelo").textContent = found["MODELO"];
      document.getElementById("calibracion").textContent = found["FECHA DE CALIBRACION"];
      document.getElementById("proxima").textContent = found["FECHA PROXIMA CALIBRACIÓN"];
      document.getElementById("estado").textContent = found["ESTADO"];
      document.getElementById("rango").textContent = found["RANGO"];
    } else {
      document.getElementById("notfound").textContent = "⚠ Instrumento no encontrado.";
    }
  })
  .catch(err => {
    document.getElementById("notfound").textContent = "⚠ Error al cargar datos.";
    console.error(err);
  });
