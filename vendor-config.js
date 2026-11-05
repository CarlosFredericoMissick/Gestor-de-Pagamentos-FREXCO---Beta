// Configuração das bibliotecas carregadas via CDN (precisa rodar no <head>, antes do Tailwind processar a página).
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
tailwind.config = {
    theme: {
        extend: {
            colors: {
                primary: '#8DC63F',
                'primary-dark': '#7AB82E',
                'primary-light': '#A3D35A',
                'frexco': '#8DC63F',
                'frexco-dark': '#7AB82E',
            }
        }
    }
}
