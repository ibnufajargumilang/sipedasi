// Tailwind Config
tailwind.config = {
    darkMode: 'class',
    theme: {
        extend: {
            fontFamily: {
                sans: ['Inter', 'sans-serif'],
                heading: ['Outfit', 'sans-serif'],
            },
            colors: {
                brand: {
                    50: '#f0f9ff',
                    100: '#e0f2fe',
                    500: '#0ea5e9',
                    600: '#0284c7',
                    700: '#0369a1',
                    900: '#0c4a6e',
                }
            }
        }
    }
}

// Theme Logic (Synced with Dashboard)
const html = document.documentElement;

// Elements
const themeToggleDesktop = document.getElementById('themeToggleDesktop');
const themeToggleMobile = document.getElementById('themeToggleMobile');

function getTheme() {
    return localStorage.getItem('sipedasi:theme') || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
}

function setTheme(theme) {
    if (theme === 'dark') {
        html.classList.add('dark');
        localStorage.setItem('sipedasi:theme', 'dark');
    } else {
        html.classList.remove('dark');
        localStorage.setItem('sipedasi:theme', 'light');
    }
    updateToggleIcons(theme);
}

function updateToggleIcons(theme) {
    const isDark = theme === 'dark';

    // Desktop Toggle Animation
    if (themeToggleDesktop) {
        // We might want to animate the toggle thumb here if we had more complex logic, 
        // but CSS classes handled in HTML structure (absolute positioning) might need adjustment if we want true animation.
        // For now, let's just toggle a class or simpler logic if needed. 
        // Actually, let's implement a simple class toggle on the thumb wrapper if we want.
        // However, a simpler way is just to toggle the global state which CSS usually reacts to.
        // Let's assume the CSS handles `dark` class on HTML.

        // Optional: Update text
        const textSpan = themeToggleDesktop.querySelector('span');
        if (textSpan) textSpan.innerHTML = isDark ? '<i class="bi bi-moon-stars-fill"></i> Mode Gelap' : '<i class="bi bi-sun-fill"></i> Mode Terang';
    }

    // Mobile Icon
    if (themeToggleMobile) {
        const icon = themeToggleMobile.querySelector('i');
        if (isDark) {
            icon.className = 'bi bi-moon-stars-fill text-xl text-yellow-400';
        } else {
            icon.className = 'bi bi-sun-fill text-xl text-orange-400';
        }
    }
}

// Initialize
setTheme(getTheme());

// Event Listeners
if (themeToggleDesktop) {
    themeToggleDesktop.addEventListener('click', () => {
        const current = getTheme();
        setTheme(current === 'dark' ? 'light' : 'dark');
    });
}

if (themeToggleMobile) {
    themeToggleMobile.addEventListener('click', () => {
        const current = getTheme();
        setTheme(current === 'dark' ? 'light' : 'dark');
    });
}

// Sidebar Active State Logic
const sections = document.querySelectorAll('section');
const navLinks = document.querySelectorAll('.sidebar-link');

window.addEventListener('scroll', () => {
    let current = '';
    sections.forEach(section => {
        const sectionTop = section.offsetTop;
        const sectionHeight = section.clientHeight;
        if (pageYOffset >= (sectionTop - 150)) {
            current = section.getAttribute('id');
        }
    });

    navLinks.forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('href').includes(current)) {
            link.classList.add('active');
        }
    });
});

// Current Year
const footerP = document.querySelector('footer p');
if (footerP) {
    footerP.innerHTML = `&copy; ${new Date().getFullYear()} Badan Pendapatan Daerah. SiRekon System v2.0`;
}

// Mobile Menu
const mobileMenuBtn = document.getElementById('mobileMenuBtn');
const sidebar = document.querySelector('aside');

if (mobileMenuBtn) {
    mobileMenuBtn.addEventListener('click', () => {
        sidebar.classList.toggle('hidden');
        sidebar.classList.toggle('fixed');
        sidebar.classList.toggle('inset-0');
        sidebar.classList.toggle('w-full');
        sidebar.classList.toggle('bg-white');
        sidebar.classList.toggle('z-40');
        sidebar.classList.toggle('pt-16'); // Add padding for header
    });
}

// Close sidebar on link click (mobile)
document.querySelectorAll('.sidebar-link').forEach(link => {
    link.addEventListener('click', () => {
        if (window.innerWidth < 768) {
            sidebar.classList.add('hidden');
            sidebar.classList.remove('fixed', 'inset-0', 'w-full', 'z-40', 'pt-16');
        }
    });
});
