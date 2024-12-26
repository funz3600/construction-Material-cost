function login() {
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;

    // Replace with your actual authentication logic
    if (username === 'admin' && password === 'password') {
        window.location.href = 'materials.html'; // Redirect to the materials page
        return false;
    } else {
        alert('Invalid username or password');
        return false;
    }
}