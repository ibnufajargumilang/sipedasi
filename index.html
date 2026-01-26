<!DOCTYPE html>
<html lang="id">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sipedasi Portal</title>

    <!-- Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&display=swap" rel="stylesheet">

    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>

    <!-- Custom Config & Styles -->
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Outfit', 'sans-serif'],
                    },
                    colors: {
                        brand: {
                            50: '#f0f9ff',
                            100: '#e0f2fe',
                            400: '#38bdf8',
                            500: '#0ea5e9',
                            600: '#0284c7',
                            900: '#0c4a6e',
                        },
                        dark: {
                            800: '#1e1e2e',
                            900: '#11111b',
                        }
                    },
                    animation: {
                        'float': 'float 6s ease-in-out infinite',
                        'blob': 'blob 7s infinite',
                    },
                    keyframes: {
                        float: {
                            '0%, 100%': {
                                transform: 'translateY(0)'
                            },
                            '50%': {
                                transform: 'translateY(-20px)'
                            },
                        },
                        blob: {
                            '0%': {
                                transform: 'translate(0px, 0px) scale(1)'
                            },
                            '33%': {
                                transform: 'translate(30px, -50px) scale(1.1)'
                            },
                            '66%': {
                                transform: 'translate(-20px, 20px) scale(0.9)'
                            },
                            '100%': {
                                transform: 'translate(0px, 0px) scale(1)'
                            },
                        }
                    }
                }
            }
        }
    </script>

    <style type="text/tailwindcss">
        @layer utilities {
            .glass {
                @apply bg-white/10 backdrop-blur-lg border border-white/20 shadow-xl;
            }
            .glass-hover {
                @apply hover:bg-white/20 hover:border-white/30 transition-all duration-300;
            }
        }
        
        body {
            background-color: #0f172a;
            color: #ffffff;
            overflow-x: hidden;
        }

        .bg-gradient-mesh {
            background-color: hsla(220, 60%, 10%, 1);
            background-image: 
                radial-gradient(at 0% 0%, hsla(253,16%,7%,1) 0, transparent 50%), 
                radial-gradient(at 50% 0%, hsla(225,39%,30%,1) 0, transparent 50%), 
                radial-gradient(at 100% 0%, hsla(339,49%,30%,1) 0, transparent 50%);
        }
    </style>
</head>

<body class="min-h-screen flex flex-col items-center justify-center relative selection:bg-brand-500 selection:text-white">

    <!-- Background Elements -->
    <div class="fixed inset-0 z-0 bg-gradient-mesh">
        <div class="absolute top-0 -left-4 w-72 h-72 bg-purple-500 rounded-full mix-blend-multiply filter blur-3xl opacity-20 animate-blob"></div>
        <div class="absolute top-0 -right-4 w-72 h-72 bg-blue-500 rounded-full mix-blend-multiply filter blur-3xl opacity-20 animate-blob animation-delay-2000"></div>
        <div class="absolute -bottom-8 left-20 w-72 h-72 bg-pink-500 rounded-full mix-blend-multiply filter blur-3xl opacity-20 animate-blob animation-delay-4000"></div>
    </div>

    <!-- Main Content -->
    <main class="z-10 w-full max-w-5xl px-4 py-16 text-center">

        <!-- Header -->
        <div class="mb-16 space-y-4 animate-float">
            <div class="inline-flex items-center justify-center p-3 mb-6 bg-white/5 rounded-2xl glass ring-1 ring-white/10">
                <svg class="w-10 h-10 text-brand-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" />
                </svg>
            </div>
            <h1 class="text-5xl md:text-7xl font-bold tracking-tight bg-clip-text text-transparent bg-gradient-to-r from-white via-blue-100 to-blue-200 drop-shadow-sm">
                SIPEDASI
            </h1>
            <p class="text-lg text-blue-200/80 font-light max-w-2xl mx-auto leading-relaxed">
                Sistem Informasi Pengelolaan Data dan Sistem Informasi
                <br>
                <span class="text-sm uppercase tracking-widest opacity-60">Select System Version</span>
            </p>
        </div>

        <!-- Version Grid -->
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 max-w-4xl mx-auto">
            <?php
            // Scan directory
            $dirs = array_filter(glob('*'), 'is_dir');
            // Exclude common system folders if any not caught by glob
            $exclude = ['.', '..', '.git', '.gemini', '.agent', 'vendor', 'node_modules'];

            $hasVersions = false;

            foreach ($dirs as $dir) {
                if (in_array($dir, $exclude)) continue;
                $hasVersions = true;

                // Determine icon based on version name (just for flavor)
                $isV1 = stripos($dir, 'v1') !== false;
                $isV2 = stripos($dir, 'v2') !== false;

                // Formatted Name
                $displayName = strtoupper($dir);
                $tagLine = "Legacy System";
                $tagColor = "bg-gray-500/20 text-gray-300";

                if ($isV2) {
                    $tagLine = "Latest Stable";
                    $tagColor = "bg-green-500/20 text-green-300";
                } elseif ($isV1) {
                    $tagLine = "Deprecated";
                    $tagColor = "bg-amber-500/20 text-amber-300";
                } else {
                    $tagLine = "Development";
                    $tagColor = "bg-blue-500/20 text-blue-300";
                }
            ?>

                <a href="<?php echo htmlspecialchars($dir); ?>/"
                    class="group relative glass rounded-3xl p-8 text-left transition-all duration-300 hover:scale-105 hover:-translate-y-2 hover:bg-white/15 border-t border-white/10 hover:border-brand-400/50 hover:shadow-brand-500/20">

                    <!-- Hover Glow Effect -->
                    <div class="absolute -inset-0.5 bg-gradient-to-r from-brand-500 to-purple-600 rounded-3xl opacity-0 group-hover:opacity-20 transition duration-500 blur-lg"></div>

                    <div class="relative z-10 flex flex-col h-full justify-between">
                        <div>
                            <div class="flex items-center justify-between mb-6">
                                <div class="p-3 bg-white/5 rounded-xl border border-white/10 group-hover:bg-brand-500/20 group-hover:border-brand-400/30 transition-colors">
                                    <svg class="w-8 h-8 text-brand-300" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M14 10l-2 1m0 0l-2-1m2 1v2.5M20 7l-2 1m2-1l-2-1m2 1v2.5M14 4l-2-1-2 1M4 7l2-1M4 7l2 1M4 7v2.5M12 21l-2-1m2 1l2-1m-2 1v-2.5M6 18l-2-1v-2.5M18 18l2-1v-2.5" />
                                    </svg>
                                </div>
                                <span class="inline-flex items-center px-3 py-1 rounded-full text-xs font-medium <?php echo $tagColor; ?> ring-1 ring-inset ring-white/10">
                                    <?php echo $tagLine; ?>
                                </span>
                            </div>

                            <h2 class="text-3xl font-bold text-white mb-2 tracking-tight group-hover:text-brand-300 transition-colors">
                                <?php echo $displayName; ?>
                            </h2>
                            <p class="text-sm text-gray-400 line-clamp-2">
                                Access the <?php echo $displayName; ?> dashboard and modules.
                            </p>
                        </div>

                        <div class="mt-8 flex items-center text-sm font-medium text-brand-300 group-hover:translate-x-2 transition-transform duration-300">
                            Enter System
                            <svg class="w-4 h-4 ml-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 8l4 4m0 0l-4 4m4-4H3" />
                            </svg>
                        </div>
                    </div>
                </a>

            <?php
            }

            if (!$hasVersions) {
            ?>
                <div class="col-span-full py-12 text-center glass rounded-3xl">
                    <p class="text-gray-400">No versions found in this directory.</p>
                </div>
            <?php } ?>

        </div>

    </main>

    <!-- Footer -->
    <footer class="absolute bottom-6 w-full text-center z-10">
        <p class="text-xs text-slate-500 uppercase tracking-widest font-semibold">
            &copy; <?php echo date('Y'); ?> Sipedasi Corp. All rights reserved.
        </p>
    </footer>

</body>

</html>