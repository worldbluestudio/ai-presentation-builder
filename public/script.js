document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('generateForm');
    const formSection = document.getElementById('formSection');
    const loadingSection = document.getElementById('loadingSection');
    const resultSection = document.getElementById('resultSection');
    const submitBtn = document.getElementById('submitBtn');
    
    // Download logic
    let pptxBlobUrl = null;
    let presentationTitle = 'presentation.pptx';

    document.getElementById('downloadBtn').addEventListener('click', () => {
        if (pptxBlobUrl) {
            const a = document.createElement('a');
            a.href = pptxBlobUrl;
            a.download = presentationTitle;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }
    });

    document.getElementById('resetBtn').addEventListener('click', () => {
        resultSection.classList.add('hidden');
        formSection.classList.remove('hidden');
        form.reset();
        
        // Clean up blob URL to free memory
        if (pptxBlobUrl) {
            URL.revokeObjectURL(pptxBlobUrl);
            pptxBlobUrl = null;
        }
    });

    // Form submission
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const prompt = document.getElementById('prompt').value;
        const theme = document.getElementById('theme').value;
        const slideCount = document.getElementById('slide_count').value;

        // Update UI
        formSection.classList.add('hidden');
        loadingSection.classList.remove('hidden');
        
        // Setup pseudo loading steps
        updateLoadingStep(1);

        try {
            // Note: Since Vercel Serverless Function might take time, we just make one big request.
            // In a production app with streaming, we would update steps based on server events.
            // Here we use a timeout simulation for visual feedback while waiting for the API.
            
            const step2Timer = setTimeout(() => updateLoadingStep(2), 15000); // Assume images generation starts around 15s
            const step3Timer = setTimeout(() => updateLoadingStep(3), 40000); // Assume building pptx around 40s

            const response = await fetch('/api/generate', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    prompt: prompt,
                    theme: theme,
                    slide_count: parseInt(slideCount)
                })
            });

            clearTimeout(step2Timer);
            clearTimeout(step3Timer);

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.error || 'Failed to generate presentation');
            }

            // Get the filename from headers if possible
            const disposition = response.headers.get('content-disposition');
            if (disposition && disposition.indexOf('filename=') !== -1) {
                const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                const matches = filenameRegex.exec(disposition);
                if (matches != null && matches[1]) {
                    presentationTitle = matches[1].replace(/['"]/g, '');
                }
            }

            // Get binary data
            const blob = await response.blob();
            pptxBlobUrl = URL.createObjectURL(blob);

            // Show Result
            loadingSection.classList.add('hidden');
            resultSection.classList.remove('hidden');

        } catch (error) {
            console.error(error);
            alert(`エラーが発生しました: ${error.message}\nVercelのタイムアウト(60秒)を超過した可能性があります。`);
            loadingSection.classList.add('hidden');
            formSection.classList.remove('hidden');
        }
    });

    function updateLoadingStep(stepNum) {
        document.querySelectorAll('.step').forEach(el => el.classList.remove('active'));
        document.getElementById(`step${stepNum}`).classList.add('active');
        
        const title = document.getElementById('loadingTitle');
        const desc = document.getElementById('loadingDesc');
        
        if (stepNum === 1) {
            title.textContent = "AIが思考しています...";
            desc.textContent = "Gemini 3.5 Flash が最適なプレゼン構成を検討しています。";
        } else if (stepNum === 2) {
            title.textContent = "画像を生成しています...";
            desc.textContent = "Nano Banana 2 が各スライドに合わせた高品質な画像を生成しています。";
        } else if (stepNum === 3) {
            title.textContent = "PowerPointを構築中...";
            desc.textContent = "スライドと画像、テキストを組み合わせて .pptx ファイルをビルドしています。";
        }
    }
});
