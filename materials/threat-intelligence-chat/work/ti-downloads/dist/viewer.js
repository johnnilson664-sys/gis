'use strict';
const image = document.getElementById('slide-image');
const number = document.getElementById('slide-number');
const title = document.getElementById('slide-title');
const previous = document.getElementById('previous');
const next = document.getElementById('next');
const thumbs = document.getElementById('thumbnails');
const status = document.getElementById('viewer-status');
let slides = [], current = 0;
function show(index) {
  if (!slides.length) return;
  current = Math.max(0, Math.min(slides.length - 1, index));
  const slide = slides[current];
  image.src = slide.src;
  image.alt = `Слайд ${slide.number}: ${slide.title}`;
  number.textContent = `${String(slide.number).padStart(2, '0')} / ${slides.length}`;
  title.textContent = slide.title;
  previous.disabled = current === 0;
  next.disabled = current === slides.length - 1;
  for (const [i, button] of [...thumbs.children].entries()) button.setAttribute('aria-current', String(i === current));
  status.textContent = '';
}
previous.addEventListener('click', () => show(current - 1));
next.addEventListener('click', () => show(current + 1));
document.addEventListener('keydown', event => {
  if (['INPUT', 'TEXTAREA', 'SELECT'].includes(event.target.tagName)) return;
  if (event.key === 'ArrowLeft') { event.preventDefault(); show(current - 1); }
  if (event.key === 'ArrowRight') { event.preventDefault(); show(current + 1); }
});
image.addEventListener('error', () => { status.textContent = 'Слайд не загрузился. Презентацию можно открыть в PDF по ссылке выше.'; });
fetch('slides.json').then(response => {
  if (!response.ok) throw new Error('Slide list unavailable');
  return response.json();
}).then(data => {
  slides = data;
  for (const [i, slide] of slides.entries()) {
    const button = document.createElement('button');
    button.setAttribute('aria-label', `Слайд ${slide.number}: ${slide.title}`);
    button.setAttribute('aria-current', 'false');
    const thumb = document.createElement('img');
    thumb.src = slide.src; thumb.alt = ''; thumb.loading = 'lazy';
    const label = document.createElement('span'); label.textContent = String(slide.number).padStart(2, '0');
    button.append(thumb, label); button.addEventListener('click', () => show(i)); thumbs.append(button);
  }
  show(0);
}).catch(() => { status.textContent = 'Список слайдов не загрузился. Все файлы доступны по ссылкам выше.'; });
