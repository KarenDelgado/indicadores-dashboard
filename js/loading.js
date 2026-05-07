const loading = document.getElementById("loading");

function mostrarLoading() {
  if (loading) {
    loading.style.display = "flex";
  }
}

function ocultarLoading() {
  if (loading) {
    loading.style.display = "none";
  }
}