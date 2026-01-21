function upload() {
  const msg = document.getElementById("msg");
  msg.innerText = "Connecting...";

  fetch("/api/process", { method: "POST" })
    .then(res => res.json())
    .then(data => msg.innerText = data.message)
    .catch(() => msg.innerText = "Error");
}