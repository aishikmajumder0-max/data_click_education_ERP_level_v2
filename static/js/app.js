async function api(url, options={}) {
  const r = await fetch(url, options);
  const d = await r.json();
  if(!r.ok) throw new Error(d.error || 'Request failed');
  return d;
}
function esc(value){
  return String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));
}
function toggleSidebar(){document.getElementById('sidebar').classList.toggle('open')}
function showToast(message,type='success'){
  const t=document.getElementById('toast');
  if(!t){alert(message);return;}
  t.textContent=message;t.className='toast show '+type;
  setTimeout(()=>t.className='toast',2800);
}
