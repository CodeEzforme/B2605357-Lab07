document.querySelectorAll('a[href^="#"]').forEach((link)=>{
  link.addEventListener('click',()=>link.blur());
});
