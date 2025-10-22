document.addEventListener('DOMContentLoaded', () => {
    const excelFileInput = document.getElementById('excelFileInput');
    const generateCardsBtn = document.getElementById('generateCardsBtn');
    const reviewCardsContainer = document.getElementById('reviewCardsContainer');
    const reviewCardTemplate = document.getElementById('reviewCardTemplate');

    let studentData = [];

    excelFileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (file) {
            readExcelFile(file);
            generateCardsBtn.disabled = false;
        } else {
            generateCardsBtn.disabled = true;
            reviewCardsContainer.innerHTML = ''; // Clear existing cards
            studentData = [];
        }
    });

    generateCardsBtn.addEventListener('click', () => {
        if (studentData.length > 0) {
            generateReviewCards(studentData);
            downloadAllCardsAsPng();
        } else {
            alert("Please upload an Excel file first.");
        }
    });

    function readExcelFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            studentData = XLSX.utils.sheet_to_json(worksheet);

            // Basic validation for required columns
            const requiredColumns = ['Name', 'School', 'Class', 'Rating', 'Feedback', 'Image Link'];
            const missingColumns = requiredColumns.filter(col => !studentData[0] || !studentData[0].hasOwnProperty(col));

            if (missingColumns.length > 0) {
                alert(`Error: Missing required columns in your Excel file: ${missingColumns.join(', ')}. 
                       Please ensure your Excel file has columns named 'Name', 'School', 'Class', 'Rating', 'Feedback', and 'Image Link'.`);
                studentData = []; // Clear data if columns are missing
                generateCardsBtn.disabled = true;
            } else {
                console.log('Parsed Student Data:', studentData);
                // Optionally display cards immediately after parsing
                // generateReviewCards(studentData);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function generateReviewCards(data) {
        reviewCardsContainer.innerHTML = ''; // Clear previous cards
        data.forEach((student, index) => {
            const card = reviewCardTemplate.content.cloneNode(true).firstElementChild;

            card.querySelector('.profile-pic').src = student['Image Link'] || 'https://via.placeholder.com/120/E0E0E0/808080?text=No+Pic'; // Placeholder if no image link
            card.querySelector('.profile-pic').alt = student['Name'] + ' Profile';
            card.querySelector('.feedback-text').textContent = student['Feedback'];
            card.querySelector('.student-name').textContent = student['Name'];
            card.querySelector('.student-school').textContent = student['School'];
            card.querySelector('.student-class').textContent = student['Class'];

            // Generate star rating
            const starRatingContainer = card.querySelector('.star-rating');
            starRatingContainer.innerHTML = ''; // Clear existing stars
            const rating = parseInt(student['Rating']);
            for (let i = 0; i < 5; i++) {
                const star = document.createElement('span');
                star.classList.add('star');
                star.innerHTML = i < rating ? '&#9733;' : '&#9734;'; // Filled star or empty star
                starRatingContainer.appendChild(star);
            }

            // Set a unique ID for each card for easier screenshotting
            card.id = `reviewCard_${index}`;
            reviewCardsContainer.appendChild(card);
        });
    }

    async function downloadAllCardsAsPng() {
        if (studentData.length === 0) {
            alert("No review cards to download.");
            return;
        }

        generateCardsBtn.textContent = 'Generating... Please wait!';
        generateCardsBtn.disabled = true;

        for (let i = 0; i < studentData.length; i++) {
            const cardElement = document.getElementById(`reviewCard_${i}`);
            if (cardElement) {
                // Temporarily make the card visible and at full resolution for screenshot
                cardElement.style.position = 'absolute';
                cardElement.style.left = '-9999px';
                cardElement.style.top = '-9999px';
                cardElement.style.width = '400px'; // Set to a fixed width for consistent export
                cardElement.style.maxWidth = '400px';
                cardElement.style.transform = 'scale(1)'; // Ensure no scaling issues

                await html2canvas(cardElement, {
                    scale: 2, // Capture at a higher resolution (e.g., 2x)
                    useCORS: true, // Important for images from Google Drive
                    allowTaint: true, // Allow images from other origins (though useCORS is better)
                    backgroundColor: null, // Transparent background if card itself has transparency
                }).then(canvas => {
                    const link = document.createElement('a');
                    link.download = `${studentData[i]['Name'].replace(/[^a-zA-Z0-9]/g, '_')}_ReviewCard.png`;
                    link.href = canvas.toDataURL('image/png');
                    link.click();
                }).catch(error => {
                    console.error(`Error generating image for ${studentData[i]['Name']}:`, error);
                    alert(`Failed to generate image for ${studentData[i]['Name']}. Please check the console for details.`);
                });

                // Revert styles after screenshot
                cardElement.style.position = '';
                cardElement.style.left = '';
                cardElement.style.top = '';
                cardElement.style.width = '';
                cardElement.style.maxWidth = '';
                cardElement.style.transform = '';
            }
        }
        alert("All review cards have been generated and downloaded!");
        generateCardsBtn.textContent = 'Generate & Download Review Cards';
        generateCardsBtn.disabled = false;
    }
});
