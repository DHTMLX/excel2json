export default {
  input: "js/worker.js",
  output:{
  	format: "iife",
  	file: "index.js"
  },
  moduleContext:() => "self"
};