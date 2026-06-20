/* Interaction Script for 「真靈光」Global Mind-Body-Spirit Ecosystem Proposal Web */

document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const body = document.body;
    const themeToggleBtn = document.getElementById('theme-toggle');
    const menuToggleBtn = document.getElementById('menu-toggle');
    const sidebar = document.getElementById('sidebar');
    const navLinks = document.querySelectorAll('.nav-link');
    const sections = document.querySelectorAll('.main-content .section');
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabPanels = document.querySelectorAll('.tab-panel');

    let growthChartInstance = null;
    let chartRevealed = false;

    // ==========================================
    // 1. Mobile Sidebar Menu Toggle
    // ==========================================
    if (menuToggleBtn && sidebar) {
        menuToggleBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            sidebar.classList.toggle('show');
            // Toggle hamburger icon
            const icon = menuToggleBtn.querySelector('i');
            if (sidebar.classList.contains('show')) {
                icon.className = 'fa-solid fa-xmark';
            } else {
                icon.className = 'fa-solid fa-bars';
            }
        });

        // Close sidebar when clicking outside on mobile
        document.addEventListener('click', (e) => {
            if (window.innerWidth <= 768 && sidebar.classList.contains('show') && !sidebar.contains(e.target) && e.target !== menuToggleBtn) {
                sidebar.classList.remove('show');
                menuToggleBtn.querySelector('i').className = 'fa-solid fa-bars';
            }
        });

        // Close sidebar when clicking on a nav link on mobile
        navLinks.forEach(link => {
            link.addEventListener('click', () => {
                if (window.innerWidth <= 768) {
                    sidebar.classList.remove('show');
                    menuToggleBtn.querySelector('i').className = 'fa-solid fa-bars';
                }
            });
        });
    }

    // ==========================================
    // 2. Tab Switching (Financial Matrix)
    // ==========================================
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.getAttribute('aria-controls');

            // Deactivate all tabs
            tabBtns.forEach(b => {
                b.classList.remove('active');
                b.setAttribute('aria-selected', 'false');
            });
            tabPanels.forEach(p => p.classList.remove('active'));

            // Activate current tab
            btn.classList.add('active');
            btn.setAttribute('aria-selected', 'true');
            const activePanel = document.getElementById(targetId);
            if (activePanel) {
                activePanel.classList.add('active');
            }
        });
    });

    // ==========================================
    // 3. Scrollspy (IntersectionObserver)
    // ==========================================
    const scrollspyOptions = {
        root: null,
        rootMargin: '0px -20% -60% 0px', // Trigger when section occupies center-upper view
        threshold: 0.1
    };

    const scrollspyCallback = (entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const activeId = entry.target.getAttribute('id');
                navLinks.forEach(link => {
                    if (link.getAttribute('href') === `#${activeId}`) {
                        link.classList.add('active');
                    } else {
                        link.classList.remove('active');
                    }
                });
            }
        });
    };

    const observer = new IntersectionObserver(scrollspyCallback, scrollspyOptions);
    sections.forEach(section => observer.observe(section));

    // Smooth scroll offset adjustment for mobile header
    navLinks.forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const targetId = link.getAttribute('href');
            const targetSection = document.querySelector(targetId);
            if (targetSection) {
                const headerOffset = window.innerWidth <= 768 ? 70 : 0;
                const elementPosition = targetSection.getBoundingClientRect().top;
                const offsetPosition = elementPosition + window.pageYOffset - headerOffset;

                window.scrollTo({
                    top: offsetPosition,
                    behavior: 'smooth'
                });
            }
        });
    });

    // ==========================================
    // Scroll Reveal Animation (IntersectionObserver)
    // ==========================================
    const revealElements = document.querySelectorAll('.reveal');
    const revealOptions = {
        root: null,
        threshold: 0.1,
        rootMargin: '0px 0px -50px 0px' // Trigger slightly before entering viewport
    };

    const revealCallback = (entries, observer) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('reveal-visible');
                
                // If the projections grid enters, trigger the chart draw animation
                if (entry.target.classList.contains('projection-grid')) {
                    chartRevealed = true;
                    initGrowthChart();
                }
                
                observer.unobserve(entry.target); // Animate only once
            }
        });
    };

    const revealObserver = new IntersectionObserver(revealCallback, revealOptions);
    revealElements.forEach(el => revealObserver.observe(el));

    // ==========================================
    // 4. Chart.js Growth Chart Initialization
    // ==========================================
    const initGrowthChart = () => {
        const ctx = document.getElementById('growthChart');
        if (!ctx) return;
        if (!chartRevealed) return;

        const isDarkMode = body.classList.contains('dark-mode');
        
        // Dynamic theme-based colors
        const textColor = isDarkMode ? '#B3A0A3' : '#666666';
        const gridColor = isDarkMode ? 'rgba(46, 23, 25, 0.5)' : 'rgba(235, 235, 235, 0.8)';
        const goldColor = '#D4AF37';
        const redColor = isDarkMode ? '#FF5A60' : '#9C1F26';

        // Data: 3 Years Revenue and Net Profit (in Millions NTD)
        const years = ['第一年 (封測打底)', '第二年 (加盟大招商)', '第三年 (大顯化期)'];
        const revenues = [9.597, 350.0, 3500.0];    // 百萬 NTD
        const profits = [1.2, 35.0, 350.0];         // 百萬 NTD

        const chartConfig = {
            type: 'line',
            data: {
                labels: years,
                datasets: [
                    {
                        label: '年度總營業額 (百萬元)',
                        data: revenues,
                        borderColor: redColor,
                        backgroundColor: isDarkMode ? 'rgba(255, 90, 96, 0.1)' : 'rgba(156, 31, 38, 0.05)',
                        borderWidth: 3,
                        pointBackgroundColor: redColor,
                        pointBorderColor: '#fff',
                        pointHoverRadius: 8,
                        pointRadius: 6,
                        tension: 0.3,
                        yAxisID: 'y'
                    },
                    {
                        label: '集團整體純淨利 (百萬元)',
                        data: profits,
                        borderColor: goldColor,
                        backgroundColor: 'rgba(212, 175, 55, 0.05)',
                        borderWidth: 3,
                        pointBackgroundColor: goldColor,
                        pointBorderColor: '#fff',
                        pointHoverRadius: 8,
                        pointRadius: 6,
                        tension: 0.3,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,

                interaction: {
                    mode: 'index',
                    intersect: false,
                },
                plugins: {
                    legend: {
                        position: 'top',
                        labels: {
                            color: textColor,
                            font: {
                                family: 'Inter',
                                size: 18,
                                weight: '500'
                            }
                        }
                    },
                    tooltip: {
                        backgroundColor: isDarkMode ? '#1E1012' : '#ffffff',
                        titleColor: isDarkMode ? '#EAEAEA' : '#2A2A2A',
                        bodyColor: isDarkMode ? '#B3A0A3' : '#666666',
                        borderColor: goldColor,
                        borderWidth: 1,
                        padding: 12,
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed.y !== null) {
                                    label += new Intl.NumberFormat('zh-TW', { style: 'currency', currency: 'TWD', maximumFractionDigits: 1 }).format(context.parsed.y * 1000000);
                                }
                                return label;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        grid: {
                            color: gridColor
                        },
                        ticks: {
                            color: textColor,
                            font: {
                                family: 'Inter',
                                size: 16
                            }
                        }
                    },
                    y: {
                        type: 'linear',
                        display: true,
                        position: 'left',
                        title: {
                            display: true,
                            text: '營業額 (百萬元)',
                            color: textColor,
                            font: {
                                family: 'Inter',
                                size: 18,
                                weight: '600'
                            }
                        },
                        grid: {
                            color: gridColor
                        },
                        ticks: {
                            color: textColor,
                            font: {
                                family: 'Inter',
                                size: 16
                            }
                        }
                    },
                    y1: {
                        type: 'linear',
                        display: true,
                        position: 'right',
                        title: {
                            display: true,
                            text: '純淨利 (百萬元)',
                            color: textColor,
                            font: {
                                family: 'Inter',
                                size: 18,
                                weight: '600'
                            }
                        },
                        grid: {
                            drawOnChartArea: false, // only want the grid lines for one axis
                        },
                        ticks: {
                            color: textColor,
                            font: {
                                family: 'Inter',
                                size: 16
                            }
                        }
                    }
                }
            }
        };

        if (growthChartInstance) {
            growthChartInstance.destroy();
        }
        growthChartInstance = new Chart(ctx, chartConfig);
    };

    // ==========================================
    // 5. Dark Mode Logic & Chart Synchronization
    // ==========================================
    const setTheme = (isDark) => {
        const icon = themeToggleBtn.querySelector('i');
        if (isDark) {
            body.classList.remove('light-mode');
            body.classList.add('dark-mode');
            icon.className = 'fa-solid fa-sun';
            localStorage.setItem('theme', 'dark');
        } else {
            body.classList.remove('dark-mode');
            body.classList.add('light-mode');
            icon.className = 'fa-solid fa-moon';
            localStorage.setItem('theme', 'light');
        }
        // Re-initialize chart to update colors
        initGrowthChart();
    };

    // Theme Toggle Click Handler
    if (themeToggleBtn) {
        themeToggleBtn.addEventListener('click', () => {
            const isDark = body.classList.contains('dark-mode');
            setTheme(!isDark);
        });
    }

    // LocalStorage / System Preference Detection
    const savedTheme = localStorage.getItem('theme');
    const systemPrefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;

    if (savedTheme === 'dark' || (!savedTheme && systemPrefersDark)) {
        setTheme(true);
    } else {
        setTheme(false);
    }

    // ==========================================
    // 6. Taiwan Map Interactive Highlights
    // ==========================================
    const mapPins = document.querySelectorAll('.map-pin');
    const storeCards = {
        'taipei': document.getElementById('card-taipei'),
        'franchise': document.getElementById('card-franchise'),
        'village': document.getElementById('card-franchise')
    };

    mapPins.forEach(pin => {
        pin.addEventListener('mouseenter', () => {
            const storeType = pin.getAttribute('data-store');
            const targetCard = storeCards[storeType];
            if (targetCard) {
                targetCard.style.transform = 'translateY(-5px) scale(1.01)';
                targetCard.style.boxShadow = '0 12px 30px var(--color-shadow-hover)';
                targetCard.style.borderColor = 'var(--color-secondary)';
            }
        });

        pin.addEventListener('mouseleave', () => {
            const storeType = pin.getAttribute('data-store');
            const targetCard = storeCards[storeType];
            if (targetCard) {
                targetCard.style.transform = '';
                targetCard.style.boxShadow = '';
                targetCard.style.borderColor = '';
            }
        });
    });
});
