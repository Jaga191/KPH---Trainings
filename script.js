let users = JSON.parse(localStorage.getItem('users')) || {};
let accountData = JSON.parse(localStorage.getItem('accountData')) || [];
let enquiryData = JSON.parse(localStorage.getItem('enquiryData')) || [];
let demoData = JSON.parse(localStorage.getItem('demoData')) || [];
let studentData = JSON.parse(localStorage.getItem('studentData')) || [];
let isFirstLogin = JSON.parse(localStorage.getItem('firstLogin')) || {};
let userLastEnquiryNumbers = JSON.parse(localStorage.getItem('userLastEnquiryNumbers')) || {};
const DEFAULT_USERNAME = 'ryepuri';
const DEFAULT_PASSWORD = '123456789';
let currentUserEmail = null;
let forgotPasswordType = null;
const START_ENQUIRY = 1600;

// Start fresh today (June 18, 2025)
const today = new Date('2025-06-18').toISOString().split('T')[0];
if (!localStorage.getItem('enquiryDataReset') || localStorage.getItem('enquiryDataReset') !== today) {
    enquiryData = [];
    demoData = [];
    studentData = [];
    localStorage.setItem('enquiryData', JSON.stringify(enquiryData));
    localStorage.setItem('demoData', JSON.stringify(demoData));
    localStorage.setItem('studentData', JSON.stringify(studentData));
    localStorage.setItem('enquiryDataReset', today);
    // Reset userLastEnquiryNumbers for new day, starting from START_ENQUIRY
    Object.keys(userLastEnquiryNumbers).forEach(email => {
        userLastEnquiryNumbers[email] = START_ENQUIRY;
    });
    if (!userLastEnquiryNumbers[currentUserEmail]) {
        userLastEnquiryNumbers[currentUserEmail] = START_ENQUIRY;
    }
    saveStoredData();
}

// Initialize default user
if (!users[DEFAULT_USERNAME]) {
    users[DEFAULT_USERNAME] = {
        password: DEFAULT_PASSWORD,
        firstName: 'ryepuri',
        lastName: '',
        dob: '01-05-1990',
        email: 'admin@example.com',
        gender: 'Other',
        education: 'Graduation',
        maritalStatus: 'Single',
        access: { enquiry: true, demo: true, student: true }
    };
}

// Initialize workbook for data storage
let workbook = null;
function initializeWorkbook() {
    const storedWorkbook = localStorage.getItem('workbook');
    try {
        if (storedWorkbook) {
            const workbookArray = new Uint8Array(atob(storedWorkbook).split('').map(char => char.charCodeAt(0)));
            workbook = XLSX.read(workbookArray, { type: 'array' });
        } else {
            throw new Error('No stored workbook');
        }
    } catch (e) {
        workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([]), 'Accounts');
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([]), 'Enquiry');
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([]), 'Demo');
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([]), 'Student');
    }
    workbook.Sheets['Enquiry'] = XLSX.utils.json_to_sheet([]);
    saveWorkbookToStorage();
}
initializeWorkbook();

function saveWorkbookToStorage() {
    try {
        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const base64String = btoa(String.fromCharCode(...new Uint8Array(wbout)));
        localStorage.setItem('workbook', base64String);
    } catch (e) {
        console.error('Failed to save workbook:', e);
    }
}

function saveStoredData() {
    try {
        localStorage.setItem('users', JSON.stringify(users));
        localStorage.setItem('accountData', JSON.stringify(accountData));
        localStorage.setItem('enquiryData', JSON.stringify(enquiryData));
        localStorage.setItem('demoData', JSON.stringify(demoData));
        localStorage.setItem('studentData', JSON.stringify(studentData));
        localStorage.setItem('firstLogin', JSON.stringify(isFirstLogin));
        localStorage.setItem('userLastEnquiryNumbers', JSON.stringify(userLastEnquiryNumbers));
    } catch (e) {
        console.error('Failed to save data:', e);
    }
}

function formatDate(dateStr) {
    if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;
    const [year, month, day] = dateStr.split('-');
    return `${day}-${month}-${year}`;
}

function validateEmail(emailInput) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(emailInput);
}

function parseFormattedNumber(value) {
    if (!value) return 0;
    return parseFloat(value.replace(/,/g, '')) || 0;
}

function formatNumber(value) {
    return value.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function togglePassword(fieldId) {
    const passwordInput = document.getElementById(fieldId);
    const eyeIcon = passwordInput?.nextElementSibling;
    if (!passwordInput || !eyeIcon) return;
    if (passwordInput.type === 'password') {
        passwordInput.type = 'text';
        eyeIcon.textContent = '👁️';
    } else {
        passwordInput.type = 'password';
        eyeIcon.textContent = '👁️‍🗨️';
    }
}

function toggleCheckbox(checkboxId) {
    const checkboxElement = document.getElementById(checkboxId);
    const checkboxLabel = checkboxElement?.parentElement;
    const checkmarkIcon = checkboxLabel?.querySelector('.checkmark');
    if (!checkboxElement || !checkmarkIcon) return;
    checkmarkIcon.style.display = checkboxElement.checked ? 'inline-block' : 'none';
}

function generateEnquiryId() {
    // Get the last used Enquiry ID number for the current user, default to START_ENQUIRY if none
    const lastNumber = userLastEnquiryNumbers[currentUserEmail] || START_ENQUIRY;
    return `KPH${lastNumber}`;
}

function setEnquiryFormDate() {
    const dateField = document.getElementById('enquiryDate');
    if (dateField) {
        const currentDate = new Date().toISOString().split('T')[0];
        dateField.value = currentDate;
        dateField.setAttribute('min', currentDate);
    }
}

function populateEnquiryIds() {
    const demoEnquirySelect = document.getElementById('demoEnquiryId');
    const studentEnquirySelect = document.getElementById('studentEnquiryId');
    if (demoEnquirySelect) {
        demoEnquirySelect.innerHTML = '<option value="" disabled selected>Enquiry ID</option>';
        enquiryData.forEach(data => {
            const option = document.createElement('option');
            option.value = data['ENQUIRY ID'];
            option.textContent = data['ENQUIRY ID'];
            demoEnquirySelect.appendChild(option);
        });
    }
    if (studentEnquirySelect) {
        studentEnquirySelect.innerHTML = '<option value="" disabled selected>Enquiry ID</option>';
        enquiryData.forEach(data => {
            const option = document.createElement('option');
            option.value = data['ENQUIRY ID'];
            option.textContent = data['ENQUIRY ID'];
            studentEnquirySelect.appendChild(option);
        });
    }
}

function setDemoDateConstraints() {
    const demoDateField = document.getElementById('demoDate');
    if (demoDateField) {
        const currentDate = new Date().toISOString().split('T')[0];
        demoDateField.setAttribute('min', currentDate);
    }
}

function autoFillDemoForm() {
    const selectedEnquiryId = document.getElementById('demoEnquiryId')?.value;
    const enquiryRecord = enquiryData.find(data => data['ENQUIRY ID'] === selectedEnquiryId);
    const fullNameField = document.getElementById('demoFullName');
    const countryCodeField = document.getElementById('demoCountryCode');
    const phoneField = document.getElementById('demoPhone');
    const emailField = document.getElementById('demoEmail');
    const subjectField = document.getElementById('subject');

    if (enquiryRecord && fullNameField && countryCodeField && phoneField && emailField && subjectField) {
        fullNameField.value = enquiryRecord['FULL NAME'] || '';
        countryCodeField.value = enquiryRecord['COUNTRY CODE'] || '+91';
        phoneField.value = enquiryRecord['PHONE NUMBER'] || '';
        emailField.value = enquiryRecord['EMAIL ID'] || '';
        subjectField.value = enquiryRecord['COURSE OF ENQUIRY'] || '';
    } else {
        if (fullNameField) fullNameField.value = '';
        if (countryCodeField) countryCodeField.value = '+91';
        if (phoneField) phoneField.value = '';
        if (emailField) emailField.value = '';
        if (subjectField) subjectField.value = '';
    }
}

function autoFillStudentForm() {
    const selectedEnquiryId = document.getElementById('studentEnquiryId')?.value;
    const enquiryRecord = enquiryData.find(data => data['ENQUIRY ID'] === selectedEnquiryId);
    const fullNameField = document.getElementById('studentFullName');
    const countryCodeField = document.getElementById('studentCountryCode');
    const phoneField = document.getElementById('studentPhone');
    const emailField = document.getElementById('studentEmail');
    const subjectField = document.getElementById('studentSubject');

    if (enquiryRecord && fullNameField && countryCodeField && phoneField && emailField && subjectField) {
        fullNameField.value = enquiryRecord['FULL NAME'] || '';
        countryCodeField.value = enquiryRecord['COUNTRY CODE'] || '+91';
        phoneField.value = enquiryRecord['PHONE NUMBER'] || '';
        emailField.value = enquiryRecord['EMAIL ID'] || '';
        subjectField.value = enquiryRecord['COURSE OF ENQUIRY'] || '';
    } else {
        if (fullNameField) fullNameField.value = '';
        if (countryCodeField) countryCodeField.value = '+91';
        if (phoneField) phoneField.value = '';
        if (emailField) emailField.value = '';
        if (subjectField) subjectField.value = '';
    }
}

function showSection(sectionId) {
    const sections = [
        'firstLoginSection', 'defaultLoginSection', 'forgotPasswordSection',
        'createAccountSection', 'changePasswordSection', 'formsSection',
        'enquiryForm', 'demoForm', 'studentInfoForm'
    ];
    sections.forEach(id => {
        const section = document.getElementById(id);
        if (section) section.style.display = id === sectionId ? 'block' : 'none';
    });
}

function showFirstLogin() {
    showSection('firstLoginSection');
}

function showDefaultLogin() {
    showSection('defaultLoginSection');
}

function showForgotPassword(loginType) {
    forgotPasswordType = loginType;
    showSection('forgotPasswordSection');
    const emailField = document.getElementById('forgotEmail');
    if (emailField) emailField.placeholder = loginType === 'first' ? 'Email ID' : 'Username';
}

function showCreateAccount() {
    showSection('createAccountSection');
}

function showChangePassword() {
    showSection('changePasswordSection');
}

function showFormsPage() {
    showSection('formsSection');
    const formButtonsContainer = document.getElementById('formButtons');
    if (formButtonsContainer) {
        formButtonsContainer.innerHTML = '';
        const userAccess = users[currentUserEmail]?.access || {};
        if (userAccess.enquiry) {
            formButtonsContainer.innerHTML += '<button class="signup-btn form-btn" onclick="showEnquiryForm()">Enquiry Form</button>';
        }
        if (userAccess.demo) {
            formButtonsContainer.innerHTML += '<button class="signup-btn form-btn" onclick="showDemoForm()">Demo Form</button>';
        }
        if (userAccess.student) {
            formButtonsContainer.innerHTML += '<button class="signup-btn form-btn" onclick="showStudentInfoForm()">Student Information Form</button>';
        }
    }
}

function showEnquiryForm() {
    showSection('enquiryForm');
    const enquiryIdField = document.getElementById('enquiryId');
    if (enquiryIdField) enquiryIdField.value = generateEnquiryId();
    setEnquiryFormDate();
}

function showDemoForm() {
    showSection('demoForm');
    populateEnquiryIds();
    setDemoDateConstraints();
}

function showStudentInfoForm() {
    showSection('studentInfoForm');
    populateEnquiryIds();
}

function goBack() {
    showFormsPage();
}

function generateEmail() {
    const firstName = document.getElementById('firstName')?.value?.trim().toLowerCase();
    const lastName = document.getElementById('lastName')?.value?.trim().toLowerCase();
    const emailField = document.getElementById('contact');
    if (firstName && lastName && emailField) {
        emailField.value = `${firstName}.${lastName}@gmail.com`;
    } else if (emailField) {
        emailField.value = '';
    }
}

function firstLogin() {
    const emailInput = document.getElementById('email')?.value?.trim();
    const passwordInput = document.getElementById('password')?.value;

    if (!emailInput || !passwordInput) {
        alert('Please enter both email and password.');
        return;
    }

    if (users[emailInput] && users[emailInput].password === passwordInput) {
        currentUserEmail = emailInput;
        if (isFirstLogin[emailInput]) {
            showChangePassword();
        } else {
            showFormsPage();
        }
    } else {
        alert('Invalid email or password.');
    }
}

function defaultLogin() {
    const usernameInput = document.getElementById('username')?.value?.trim();
    const passwordInput = document.getElementById('defaultPassword')?.value;

    if (!usernameInput || !passwordInput) {
        alert('Please enter both username and password.');
        return;
    }

    if (usernameInput === DEFAULT_USERNAME && users[usernameInput]?.password === passwordInput) {
        showCreateAccount();
    } else {
        alert('Invalid username or password.');
    }
}

function createAccount() {
    const firstName = document.getElementById('firstName')?.value?.trim();
    const lastName = document.getElementById('lastName')?.value?.trim();
    const dob = formatDate(document.getElementById('dob')?.value);
    const gender = document.getElementById('gender')?.value;
    const email = document.getElementById('contact')?.value?.trim();
    const education = document.getElementById('education')?.value;
    const maritalStatus = document.getElementById('maritalStatus')?.value;
    const newPassword = document.getElementById('newPassword')?.value;
    const enquiryAccess = document.getElementById('enquiryAccess')?.checked;
    const demoAccess = document.getElementById('demoAccess')?.checked;
    const studentAccess = document.getElementById('studentAccess')?.checked;

    if (!firstName || !lastName || !dob || !gender || !email || !education || !maritalStatus || !newPassword) {
        alert('Please fill out all fields.');
        return;
    }

    if (!validateEmail(email)) {
        alert('Please enter a valid email address.');
        return;
    }

    if (!/^\d{2}-\d{2}-\d{4}$/.test(dob)) {
        alert('Please select a valid date of birth.');
        return;
    }

    if (!enquiryAccess && !demoAccess && !studentAccess) {
        alert('Please select at least one access permission.');
        return;
    }

    if (newPassword.length < 8) {
        alert('Password must be at least 8 characters long.');
        return;
    }

    const isExistingUser = users[email];
    let newAccess = { enquiry: enquiryAccess, demo: demoAccess, student: studentAccess };

    users[email] = {
        password: newPassword,
        firstName,
        lastName,
        dob,
        gender,
        email,
        education,
        maritalStatus,
        access: newAccess
    };

    if (!isExistingUser) {
        isFirstLogin[email] = true;
        // Set next Enquiry ID to START_ENQUIRY for new account
        userLastEnquiryNumbers[email] = START_ENQUIRY;
    }

    const accessPermissions = `${enquiryAccess ? 'Enquiry, ' : ''}${demoAccess ? 'Demo, ' : ''}${studentAccess ? 'Student' : ''}`.replace(/, $/, '');
    const accountRecord = {
        'FIRST NAME': firstName,
        'LAST NAME': lastName,
        'DATE OF BIRTH': dob,
        'GENDER': gender,
        'EMAIL ID': email,
        'EDUCATION': education,
        'MARITAL STATUS': maritalStatus,
        'PASSWORD': '********',
        'ACCESS': accessPermissions
    };

    const existingAccountIndex = accountData.findIndex(record => record['EMAIL ID'] === email);
    if (existingAccountIndex !== -1) {
        accountData[existingAccountIndex] = accountRecord;
    } else {
        accountData.push(accountRecord);
    }

    const accountHeaders = ['FIRST NAME', 'LAST NAME', 'DATE OF BIRTH', 'GENDER', 'EMAIL ID', 'EDUCATION', 'MARITAL STATUS', 'PASSWORD', 'ACCESS'];
    const accountWorksheet = XLSX.utils.json_to_sheet(accountData, { header: accountHeaders });
    workbook.Sheets['Accounts'] = accountWorksheet;

    saveStoredData();
    saveWorkbookToStorage();

    alert('Account created/updated successfully! Please login.');
    showFirstLogin();
    const accountForm = document.getElementById('createAccountForm');
    if (accountForm) {
        accountForm.reset();
        ['gender', 'education', 'maritalStatus'].forEach(id => {
            const select = document.getElementById(id);
            if (select) select.value = '';
        });
        ['enquiryAccess', 'demoAccess', 'studentAccess'].forEach(id => toggleCheckbox(id));
    }
}

function resetPassword() {
    const inputValue = document.getElementById('forgotEmail')?.value?.trim();
    const dobValue = formatDate(document.getElementById('forgotDob')?.value);
    const newPassword = document.getElementById('forgotNewPassword')?.value;
    const retypePassword = document.getElementById('forgotRetypePassword')?.value;

    if (!inputValue || !dobValue || !newPassword || !retypePassword) {
        alert('Please fill out all fields.');
        return;
    }

    if (newPassword !== retypePassword) {
        alert('New password and retype password do not match.');
        return;
    }

    if (newPassword.length < 8) {
        alert('New password must be at least 8 characters long.');
        return;
    }

    if (forgotPasswordType === 'first') {
        if (!validateEmail(inputValue) || !users[inputValue]) {
            alert('Email not found.');
            return;
        }
        if (users[inputValue].dob !== dobValue) {
            alert('Date of birth does not match.');
            return;
        }
        users[inputValue].password = newPassword;
        isFirstLogin[inputValue] = false;
        saveStoredData();
        alert('Password reset successfully!');
        showFirstLogin();
    } else if (forgotPasswordType === 'default') {
        if (inputValue !== DEFAULT_USERNAME || users[inputValue]?.dob !== dobValue) {
            alert('Username or date of birth does not match.');
            return;
        }
        users[inputValue].password = newPassword;
        saveStoredData();
        alert('Password reset successfully!');
        showDefaultLogin();
    }
}

function changePassword() {
    const oldPassword = document.getElementById('oldPassword')?.value;
    const newPassword = document.getElementById('newPasswordChange')?.value;
    const retypePassword = document.getElementById('retypeNewPassword')?.value;

    if (!oldPassword || !newPassword || !retypePassword) {
        alert('Please fill out all fields.');
        return;
    }

    if (!users[currentUserEmail] || users[currentUserEmail].password !== oldPassword) {
        alert('Old password is incorrect.');
        return;
    }

    if (newPassword !== retypePassword) {
        alert('New password and retype password do not match.');
        return;
    }

    if (newPassword.length < 8) {
        alert('New password must be at least 8 characters long.');
        return;
    }

    users[currentUserEmail].password = newPassword;
    isFirstLogin[currentUserEmail] = false;
    saveStoredData();
    alert('Password changed successfully!');
    showFormsPage();
}

function submitForm(formType) {
    if (formType === 'enquiryForm') {
        const enquiryId = document.getElementById('enquiryId')?.value;
        const date = formatDate(document.getElementById('enquiryDate')?.value);
        const fullName = document.getElementById('enquiryFullName')?.value?.trim();
        const countryCode = document.getElementById('countryCode')?.value;
        const phone = document.getElementById('enquiryPhone')?.value?.trim();
        const email = document.getElementById('enquiryEmail')?.value?.trim();
        const dob = formatDate(document.getElementById('enquiryDob')?.value);
        const course = document.getElementById('course')?.value?.trim();
        const source = document.getElementById('source')?.value;
        const education = document.getElementById('enquiryEducation')?.value;
        const passedOutYear = document.getElementById('passedOutYear')?.value;
        const about = document.getElementById('about')?.value;
        const mode = document.getElementById('mode')?.value;
        const batchTiming = document.getElementById('batchTiming')?.value;
        const language = document.getElementById('language')?.value;
        const demoStatus = document.getElementById('demoStatus')?.value;
        const comment = document.getElementById('comment')?.value?.trim();

        if (!enquiryId || !date || !fullName || !phone || !email || !dob || !course || !source || !education || !passedOutYear || !about || !mode || !batchTiming || !language || !demoStatus) {
            alert('Please fill out all required fields.');
            return;
        }

        if (!validateEmail(email)) {
            alert('Please enter a valid email address.');
            return;
        }

        if (!/^\d{10}$/.test(phone)) {
            alert('Phone number must be exactly 10 digits.');
            return;
        }

        const enquiryRecord = {
            'ENQUIRY ID': enquiryId,
            'DATE': date,
            'FULL NAME': fullName,
            'COUNTRY CODE': countryCode,
            'PHONE NUMBER': phone,
            'EMAIL ID': email,
            'STUDENT DATE OF BIRTH': dob,
            'COURSE OF ENQUIRY': course,
            'SOURCE OF ENQUIRY': source,
            'EDUCATION QUALIFICATION': education,
            'PASSED OUT YEAR': passedOutYear,
            'ABOUT': about,
            'MODE OF CLASSES': mode,
            'BATCH TIMINGS': batchTiming,
            'LANGUAGE': language,
            'DEMO STATUS': demoStatus,
            'COMMENT': comment
        };
        enquiryData.push(enquiryRecord);

        // Increment the user's Enquiry ID number
        userLastEnquiryNumbers[currentUserEmail] = (userLastEnquiryNumbers[currentUserEmail] || START_ENQUIRY) + 1;

        const enquiryHeaders = [
            'ENQUIRY ID', 'DATE', 'FULL NAME', 'COUNTRY CODE', 'PHONE NUMBER', 'EMAIL ID',
            'STUDENT DATE OF BIRTH', 'COURSE OF ENQUIRY', 'SOURCE OF ENQUIRY', 'EDUCATION QUALIFICATION',
            'PASSED OUT YEAR', 'ABOUT', 'MODE OF CLASSES', 'BATCH TIMINGS', 'LANGUAGE', 'DEMO STATUS', 'COMMENT'
        ];
        const enquiryWorksheet = XLSX.utils.json_to_sheet(enquiryData, { header: enquiryHeaders });
        workbook.Sheets['Enquiry'] = enquiryWorksheet;

        saveStoredData();
        saveWorkbookToStorage();
        alert('Enquiry Form submitted successfully!');
        const enquiryFormElement = document.getElementById('enquiryFormElement');
        if (enquiryFormElement) {
            enquiryFormElement.reset();
            ['countryCode', 'source', 'enquiryEducation', 'about', 'mode', 'batchTiming', 'language', 'demoStatus'].forEach(id => {
                const select = document.getElementById(id);
                if (select) select.value = '';
            });
        }
        const enquiryIdField = document.getElementById('enquiryId');
        if (enquiryIdField) enquiryIdField.value = generateEnquiryId();
        setEnquiryFormDate();
        goBack();
    } else if (formType === 'demoForm') {
        const enquiryId = document.getElementById('demoEnquiryId')?.value;
        const fullName = document.getElementById('demoFullName')?.value?.trim();
        const countryCode = document.getElementById('demoCountryCode')?.value;
        const phone = document.getElementById('demoPhone')?.value?.trim();
        const email = document.getElementById('demoEmail')?.value?.trim();
        const subject = document.getElementById('subject')?.value?.trim();
        const demoDate = formatDate(document.getElementById('demoDate')?.value);
        const tutorName = document.getElementById('tutorName')?.value?.trim();
        const demoTime = document.getElementById('demoTime')?.value;
        const demoFeedback = document.getElementById('demoFeedback')?.value;
        const enrollStatus = document.getElementById('enrollStatus')?.value;

        if (!enquiryId || !fullName || !phone || !email || !subject || !demoDate || !tutorName || !demoTime || !demoFeedback || !enrollStatus) {
            alert('Please fill out all required fields.');
            return;
        }

        if (!validateEmail(email)) {
            alert('Please enter a valid email address.');
            return;
        }

        if (!/^\d{10}$/.test(phone)) {
            alert('Phone number must be exactly 10 digits.');
            return;
        }

        const demoRecord = {
            'ENQUIRY ID': enquiryId,
            'FULL NAME': fullName,
            'COUNTRY CODE': countryCode,
            'PHONE NUMBER': phone,
            'EMAIL ID': email,
            'SUBJECT': subject,
            'DEMO DATE': demoDate,
            'TUTOR NAME': tutorName,
            'DEMO TIME': demoTime,
            'DEMO FEEDBACK': demoFeedback,
            'ENROLL STATUS': enrollStatus
        };
        demoData.push(demoRecord);

        const demoHeaders = [
            'ENQUIRY ID', 'FULL NAME', 'COUNTRY CODE', 'PHONE NUMBER', 'EMAIL ID',
            'SUBJECT', 'DEMO DATE', 'TUTOR NAME', 'DEMO TIME', 'DEMO FEEDBACK', 'ENROLL STATUS'
        ];
        const demoWorksheet = XLSX.utils.json_to_sheet(demoData, { header: demoHeaders });
        workbook.Sheets['Demo'] = demoWorksheet;

        saveStoredData();
        saveWorkbookToStorage();
        alert('Demo Form submitted successfully!');
        const demoFormElement = document.getElementById('demoFormElement');
        if (demoFormElement) {
            demoFormElement.reset();
            ['demoCountryCode', 'demoFeedback', 'enrollStatus'].forEach(id => {
                const select = document.getElementById(id);
                if (select) select.value = '';
            });
        }
        goBack();
    } else if (formType === 'studentInfoForm') {
        const enquiryId = document.getElementById('studentEnquiryId')?.value;
        const fullName = document.getElementById('studentFullName')?.value?.trim();
        const countryCode = document.getElementById('studentCountryCode')?.value;
        const phone = document.getElementById('studentPhone')?.value?.trim();
        const email = document.getElementById('studentEmail')?.value?.trim();
        const subject = document.getElementById('studentSubject')?.value?.trim();
        const totalFee = document.getElementById('totalFee')?.value;
        const paidAmount = document.getElementById('paidAmount')?.value;
        const pendingAmount = document.getElementById('pendingAmount')?.value;
        const paymentMode = document.getElementById('paymentMode')?.value;
        const trainerName = document.getElementById('trainerName')?.value?.trim();
        const comment = document.getElementById('studentComment')?.value?.trim();

        if (!enquiryId || !fullName || !phone || !email || !subject || !totalFee || !paidAmount || !pendingAmount || !paymentMode || !trainerName) {
            alert('Please fill out all required fields.');
            return;
        }

        if (!validateEmail(email)) {
            alert('Please enter a valid email address.');
            return;
        }

        if (!/^\d{10}$/.test(phone)) {
            alert('Phone number must be exactly 10 digits.');
            return;
        }

        const totalAmount = parseFormattedNumber(totalFee);
        const paidAmountValue = parseFormattedNumber(paidAmount);
        const pendingAmountValue = parseFormattedNumber(pendingAmount);

        if (Math.abs(totalAmount - paidAmountValue - pendingAmountValue) > 0.01) {
            alert('Total fee must equal paid amount plus pending amount.');
            return;
        }

        const studentRecord = {
            'ENQUIRY ID': enquiryId,
            'FULL NAME': fullName,
            'COUNTRY CODE': countryCode,
            'PHONE NUMBER': phone,
            'EMAIL ID': email,
            'SUBJECT': subject,
            'TOTAL FEE': formatNumber(totalAmount),
            'PAID AMOUNT': formatNumber(paidAmountValue),
            'PENDING AMOUNT': formatNumber(pendingAmountValue),
            'MODE OF PAYMENT': paymentMode,
            'TRAINER NAME': trainerName,
            'COMMENT': comment
        };
        studentData.push(studentRecord);

        const studentHeaders = [
            'ENQUIRY ID', 'FULL NAME', 'COUNTRY CODE', 'PHONE NUMBER', 'EMAIL ID',
            'SUBJECT', 'TOTAL FEE', 'PAID AMOUNT', 'PENDING AMOUNT', 'MODE OF PAYMENT',
            'TRAINER NAME', 'COMMENT'
        ];
        const studentWorksheet = XLSX.utils.json_to_sheet(studentData, { header: studentHeaders });
        workbook.Sheets['Student'] = studentWorksheet;

        saveStoredData();
        saveWorkbookToStorage();
        alert('Student Information Form submitted successfully!');
        const studentFormElement = document.getElementById('studentInfoFormElement');
        if (studentFormElement) {
            studentFormElement.reset();
            ['studentCountryCode', 'paymentMode'].forEach(id => {
                const select = document.getElementById(id);
                if (select) select.value = '';
            });
        }
        goBack();
    }
}

// Initialize form events
document.addEventListener('DOMContentLoaded', () => {
    const checkboxElements = document.querySelectorAll('.checkbox-label input');
    if (checkboxElements.length > 0) {
        checkboxElements.forEach(checkbox => {
            checkbox.addEventListener('change', () => toggleCheckbox(checkbox.id));
            toggleCheckbox(checkbox.id);
        });
    }

    // Phone number validation for Enquiry, Demo, and Student forms (digits only, 10 digits)
    const phoneFields = ['enquiryPhone', 'demoPhone', 'studentPhone'];
    phoneFields.forEach(fieldId => {
        const phoneField = document.getElementById(fieldId);
        if (phoneField) {
            phoneField.addEventListener('input', (e) => {
                const value = e.target.value;
                e.target.value = value.replace(/[^0-9]/g, '').slice(0, 10);
            });
        }
    });

    // Format number inputs with commas for Student Information Form
    const numberFields = ['totalFee', 'paidAmount', 'pendingAmount'];
    numberFields.forEach(fieldId => {
        const field = document.getElementById(fieldId);
        if (field) {
            field.addEventListener('input', (e) => {
                let value = e.target.value.replace(/[^0-9]/g, '');
                if (value) {
                    e.target.value = formatNumber(parseInt(value));
                } else {
                    e.target.value = '';
                }
            });
        }
    });
});