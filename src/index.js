document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM fully loaded and parsed');

    // Function to update progress for a specific console
    function updateProgress(sectionNumber, percentage, message) {
        var outputElement = document.getElementById(`output${sectionNumber}`);
        outputElement.innerHTML = `<div style="color: green;">${percentage}% completed:</div> ${message}`;
    }

    // Initialize SSE for progress updates
    var eventSource = new EventSource('/progress');
    eventSource.onmessage = function(event) {
        const data = JSON.parse(event.data);
        if (data.message && data.sectionNumber) {
            updateProgress(data.sectionNumber, data.percentage, data.message);
        }
    };

    // Function to perform data fetch for a specific section
    function performDataFetch(url, custID, sectionNumber) {
        var loadingElement = document.getElementById(`loading${sectionNumber}`);
        var outputElement = document.getElementById(`output${sectionNumber}`);
        var timerElement = document.getElementById(`timer${sectionNumber}`);

        outputElement.innerHTML = ''; // Clear previous data
        loadingElement.style.display = 'inline-block'; // Show loading animation
        timerElement.textContent = 'Loading...';

        var secondsElapsed = 0;
        var timer = setInterval(function() {
            secondsElapsed++;
            timerElement.textContent = 'Loading... ' + secondsElapsed + 's';
        }, 1000);

        fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ custID: custID }),
        })
        .then(response => {
            clearInterval(timer);
            loadingElement.style.display = 'none';

            return response.json();
        })
        .then(data => {
            // Display the result and output
            const output = data.output ? `<pre>${data.output}</pre>` : "No output returned.";
            const message = data.message ? `<p>${data.message}</p>` : "No message returned.";
            timerElement.textContent = `Total time elapsed: ${secondsElapsed}s`;
            outputElement.innerHTML = message + output;
        })
        .catch(error => {
            clearInterval(timer);
            loadingElement.style.display = 'none';
            timerElement.textContent = '';
            outputElement.innerHTML = `<div style="color: red;">Error: ${error.message}</div>`;
        });
    }

    // Function for each Section Button
    function fetchData(sectionNumber) {
        var custID = document.getElementById(`custID${sectionNumber}`).value;
        performDataFetch(`/fetchData${sectionNumber}`, custID, sectionNumber);
    }

    // Attach event listeners to buttons for all 12 sections
    for (let i = 1; i <= 12; i++) {
        document.getElementById(`fetchData${i}Btn`).addEventListener('click', function() {
            fetchData(i);
        });
    }
});
