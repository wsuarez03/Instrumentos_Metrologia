function getQueryParam(param) {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(param);
}

const id = getQueryParam("id");

fetch("instrumentos.json")
  .then(response => response.json())
  .then(data => {
    // Filtrar todos los registros que tengan el mismo IDENTIFICACIÓN
    const foundItems = data.filter(item => item["IDENTIFICACIÓN"] === id);

    if (foundItems.length > 0) {
      const container = document.getElementById("info");
      container.innerHTML = ""; // Limpiar contenido anterior

      foundItems.forEach((found, index) => {
        const block = document.createElement("div");
        block.classList.add("certificado");

        block.innerHTML = `
          <h3>Certificado ${index + 1}</h3>
          <p><strong>ID:</strong> ${found["IDENTIFICACIÓN"] ?? "-"}</p>
          <p><strong>Nombre:</strong> ${found["EQUIPO  /  INSTRUMENTO"] ?? "-"}</p>
          <p><strong>Fabricante:</strong> ${found["FABRICANTE"] ?? "-"}</p>
          <p><strong>Certificado:</strong> ${found["CERTIFICADO No. "] ?? "-"}</p>
          <p><strong>Serie:</strong> ${found["SERIE "] ?? "-"}</p>
          <p><strong>Modelo:</strong> ${found["MODELO"] ?? "-"}</p>
          <p><strong>Fecha Calibración:</strong> ${found["FECHA DE CALIBRACION"] ?? "-"}</p>
          <p><strong>Próxima Calibración:</strong> ${found["FECHA PROXIMA CALIBRACIÓN"] ?? "-"}</p>
          <p><strong>Estado:</strong> ${found["ESTADO"] ?? "-"}</p>
          <p><strong>Rango:</strong> ${found["RANGO"] ?? "-"}</p>
          <hr>
        `;

        container.appendChild(block);
      });
    } else {
      document.getElementById("notfound").textContent = "⚠ Instrumento no encontrado.";
    }
  })
  .catch(err => {
    document.getElementById("notfound").textContent = "⚠ Error al cargar datos.";
    console.error(err);
  });
