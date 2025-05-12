let questions = [];
let currentQuestionIndex = 0;
let startTime, questionStartTime;
let results = [];
let timer;

const fileUploadScreen = document.getElementById('file-upload');
const startScreen = document.getElementById('start-screen');
const bufferScreen = document.getElementById('buffer-screen');
const questionScreen = document.getElementById('question-screen');
const optionsScreen = document.getElementById('options-screen');
const questionText = document.getElementById('question-text');
const option1 = document.getElementById('option1');
const option2 = document.getElementById('option2');
const saveButton = document.getElementById('save-button');

document.getElementById('file-input').addEventListener('change', (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet);

        questions = json.map(row => ({
            word: row.Word,
            options: [row.Option1, row.Option2],
            answer: row.Answer,
            backgroundColor: row.BackgroundColor || "#FFFFFF",
            textColor: row.TextColor || "#000000"
        }));

        console.log('題目已載入:', questions);
    };

    reader.readAsArrayBuffer(file);
});

document.getElementById('next-button').addEventListener('click', () => {
    startTime = Date.now();
    fileUploadScreen.style.display = 'none';
    startScreen.style.display = 'flex';
});

startScreen.addEventListener('click', moveToBufferScreen);
document.addEventListener('keydown', (event) => {
    if (startScreen.style.display === 'flex' && event.key === 'Enter') {
        moveToBufferScreen();
    }
});

function moveToBufferScreen() {
    startScreen.style.display = 'none';
    bufferScreen.style.display = 'flex';

    setTimeout(() => {
        bufferScreen.style.display = 'none';
        startTest();
    }, 2000);
}

function startTest() {
    loadNextQuestion();
}

function loadNextQuestion() {
    if (currentQuestionIndex >= questions.length) {
        return showResults();
    }

    const currentQuestion = questions[currentQuestionIndex];
    questionText.textContent = currentQuestion.word;
    questionScreen.style.display = 'flex';

    questionStartTime = Date.now();

    setTimeout(() => {
        questionScreen.style.display = 'none';
        showOptions(currentQuestion);
    }, 2000);
}

function showOptions(currentQuestion) {
    option1.style.backgroundColor = currentQuestion.backgroundColor;
    option2.style.backgroundColor = currentQuestion.backgroundColor;
    option1.style.color = currentQuestion.textColor;
    option2.style.color = currentQuestion.textColor;

    option1.textContent = currentQuestion.options[0];
    option2.textContent = currentQuestion.options[1];

    optionsScreen.style.display = 'flex';

    timer = setTimeout(() => {
        handleAnswer(currentQuestion, "未作答");
    }, 10000);

    option1.onclick = () => handleAnswer(currentQuestion, option1.textContent);
    option2.onclick = () => handleAnswer(currentQuestion, option2.textContent);
}

document.addEventListener('keydown', (event) => {
    if (optionsScreen.style.display === 'flex') {
        if (event.key === 'ArrowLeft') {
            handleAnswer(questions[currentQuestionIndex], option1.textContent);
        } else if (event.key === 'ArrowRight') {
            handleAnswer(questions[currentQuestionIndex], option2.textContent);
        }
    }
});

function handleAnswer(currentQuestion, selectedOption) {
    clearTimeout(timer);
    const reactionTime = selectedOption === "未作答" ? "未作答" : (Date.now() - questionStartTime) / 1000;

    results.push({
        order: currentQuestionIndex + 1,
        question: currentQuestion.word,
        backgroundColor: currentQuestion.backgroundColor,
        textColor: currentQuestion.textColor,
        correctAnswer: currentQuestion.answer,
        userAnswer: selectedOption,
        reactionTime: reactionTime
    });

    optionsScreen.style.display = 'none';
    currentQuestionIndex++;
    loadNextQuestion();
}

function showResults() {
    const totalTime = (Date.now() - startTime) / 1000;
    results.push({ order: "總時間", totalTime: `${totalTime}s` });

    const worksheet = XLSX.utils.json_to_sheet(results);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Results");

    XLSX.writeFile(workbook, "test_results.xlsx");
}

saveButton.addEventListener('click', saveProgress);

function saveProgress() {
    const now = new Date();
    const timestamp = now.toISOString().replace(/[:.]/g, '-');
    const filename = `progress_${timestamp}.xlsx`;

    const worksheet = XLSX.utils.json_to_sheet(results);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Progress");

    XLSX.writeFile(workbook, filename);
}
