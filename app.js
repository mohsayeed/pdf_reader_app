const x = require("events");
const http = require("http");


const events = new x();

events.on("se",()=>{
    console.log("secon eve")
})

events.on("fi", () => {
  console.log("welocm eve");
});

events.emit("fi")

events.on("fi",()=>{
    events.emit("se")
    console.log("fir eve");
})
events.emit("fi")


const hostname = "127.0.0.1";
const port = 3000;

const server = http.createServer((req, res) => {
  res.statusCode = 200;
  res.setHeader("Content-Type", "text/plain");
  res.end("Hello World");
});

server.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});
