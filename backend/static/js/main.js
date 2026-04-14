
/* ASELBY — main.js */

// Navbar scroll
const navbar = document.getElementById('navbar');
if (navbar) {
    window.addEventListener('scroll', () => {
        navbar.classList.toggle('scrolled', window.scrollY > 60);
    });
}

// Burger menu mobile
const burger = document.getElementById('burger');
const mobileMenu = document.getElementById('mobile-menu');
if (burger && mobileMenu) {
    burger.addEventListener('click', () => {
        mobileMenu.classList.toggle('ouvert');
        burger.classList.toggle('actif');
    });
}

// Reveal on scroll
const revealEls = document.querySelectorAll('.reveal, .reveal-left, .reveal-right, .membre-carte');
if (revealEls.length) {
    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) entry.target.classList.add('visible');
        });
    }, { threshold: 0.12, rootMargin: '0px 0px -50px 0px' });
    revealEls.forEach(el => observer.observe(el));
}

// FAQ toggle
function toggleFAQ(btn) {
    const item = btn.closest('.faq-item');
    const estOuvert = item.classList.contains('ouvert');
    document.querySelectorAll('.faq-item.ouvert').forEach(i => i.classList.remove('ouvert'));
    if (!estOuvert) item.classList.add('ouvert');
}

// Barres animées
function animerBarres() {
    document.querySelectorAll('.barre-fill[data-width]').forEach(b => {
        setTimeout(() => b.style.width = b.dataset.width + '%', 300);
    });
}
animerBarres();

// Smooth scroll
document.querySelectorAll('a[href^="#"]').forEach(a => {
    a.addEventListener('click', e => {
        const cible = document.querySelector(a.getAttribute('href'));
        if (cible) { e.preventDefault(); cible.scrollIntoView({ behavior: 'smooth' }); }
    });
});

// Parallaxe hero
const heroContent = document.querySelector('.hero-content');
if (heroContent) {
    window.addEventListener('scroll', () => {
        const s = window.scrollY;
        if (s < window.innerHeight) {
            heroContent.style.transform = `translateY(${s * 0.25}px)`;
            heroContent.style.opacity = 1 - s / (window.innerHeight * 0.85);
        }
    });
}

// Date topbar dashboard
const dateEl = document.getElementById('date-auj');
if (dateEl) {
    dateEl.textContent = new Date().toLocaleDateString('fr-FR', {
        weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
    });
}
