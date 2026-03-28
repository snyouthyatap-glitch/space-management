document.addEventListener('DOMContentLoaded', () => {
    const links = document.querySelectorAll('.routine-link');
    const resetButton = document.getElementById('reset-routine');

    // Load completion status from localStorage
    const loadStatus = () => {
        const completedSteps = JSON.parse(localStorage.getItem('routine_completed') || '[]');
        completedSteps.forEach(stepId => {
            const card = document.getElementById(`step-${stepId}`);
            if (card) {
                card.classList.add('completed');
            }
        });
    };

    // Save completion status
    const saveStatus = (stepId) => {
        let completedSteps = JSON.parse(localStorage.getItem('routine_completed') || '[]');
        if (!completedSteps.includes(stepId)) {
            completedSteps.push(stepId);
        }
        localStorage.setItem('routine_completed', JSON.stringify(completedSteps));
    };

    // Handle link clicks
    links.forEach(link => {
        link.addEventListener('click', (e) => {
            const stepId = link.getAttribute('data-step');
            const card = document.getElementById(`step-${stepId}`);
            
            // Mark step as completed after a slight delay
            setTimeout(() => {
                card.classList.add('completed');
                saveStatus(stepId);
                
                // Optional: Provide visual feedback or sound
                console.log(`Step ${stepId} completed!`);
            }, 500);
        });
    });

    // Reset routine
    resetButton.addEventListener('click', () => {
        if (confirm('모든 루틴 기록을 초기화하시겠습니까?')) {
            localStorage.removeItem('routine_completed');
            document.querySelectorAll('.step-card').forEach(card => {
                card.classList.remove('completed');
            });
        }
    });

    // Initial load
    loadStatus();
});
