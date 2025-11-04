function getQueryParam(param) {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(param);
}

const id = getQueryParam("id");

fetch("instrumentos.json")
  .then(response => response.json())
  .then(data => {
    // Buscar todos los registros con el mismo ID
    const foundItems = data.filter(item => item["IDENTIFICACIÓN"] === id);

    const container = document.getElementById("info");
    const notFound = document.getElementById("notfound");

    if (foundItems.length === 0) {
      notFound.textContent = "⚠ Instrumento no encontrado.";
      return;
    }

    container.innerHTML = ""; // Limpiar contenido

    foundItems.forEach((found, index) => {
      const block = document.createElement("div");
      block.classList.add("certificado");

      // 🧱 Encabezado del certificado (si hay varios)
      const title = document.createElement("h3");
      title.textContent = `Certificado ${index + 1}`;
      block.appendChild(title);

      // 🔍 Mostrar todos los campos dinámicamente
      Object.entries(found).forEach(([clave, valor]) => {
        const p = document.createElement("p");
        p.innerHTML = `<strong>${clave}:</strong> ${valor ?? "—"}`;
        block.appendChild(p);
      });

      container.appendChild(block);
      container.appendChild(document.createElement("hr"));
    });
  })
  .catch(err => {
    document.getElementById("notfound").textContent = "⚠ Error al cargar datos.";
    console.error(err);
  });
