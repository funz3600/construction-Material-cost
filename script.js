body {
    font-family: 'Roboto', sans-serif;
    margin: 0;
    padding: 0;
    background-color: #f4f4f4;
    color: #333;
}

.hidden {
    display: none;
}

.login-page {
    width: 360px;
    padding: 8% 0 0;
    margin: auto;
}

.form {
    position: relative;
    z-index: 1;
    background: #FFFFFF;
    max-width: 360px;
    margin: 0 auto 100px;
    padding: 45px;
    text-align: center;
    box-shadow: 0 0 20px 0 rgba(0, 0, 0, 0.2), 0 5px 5px 0 rgba(0, 0, 0, 0.24);
}

.form input {
    font-family: "Roboto", sans-serif;
    outline: 0;
    background: #f2f2f2;
    width: 100%;
    border: 0;
    margin: 0 0 15px;
    padding: 15px;
    box-sizing: border-box;
    font-size: 14px;
}

.form button {
    font-family: "Roboto", sans-serif;
    text-transform: uppercase;
    outline: 0;
    background: #4CAF50;
    width: 100%;
    border: 0;
    padding: 15px;
    color: #FFFFFF;
    font-size: 14px;
    cursor: pointer;
}

.form button:hover,
.form button:active,
.form button:focus {
    background: #43A047;
}

.container {
    width: 80%;
    margin: 0 auto;
}

header {
    background-color: #333;
    color: white;
    padding: 10px 0;
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.logo img {
    height: 50px;
}

nav ul {
    list-style: none;
    padding: 0;
    margin: 0;
    display: flex;
    gap: 20px;
}

nav ul li a {
    color: white;
    text-decoration: none;
    font-weight: bold;
    padding: 10px;
    transition: background 0.3s;
}

nav ul li a:hover {
    background-color: #555;
    border-radius: 5px;
}

.materials {
    padding: 50px 0;
}

.materials h2 {
    text-align: center;
    margin-bottom: 20px;
    font-size: 2em;
    color: #444;
}

#upload, #save {
    display: block;
    margin: 0 auto 20px;
    padding: 10px;
    font-size: 1em;
    border: 1px solid #ccc;
    border-radius: 5px;
    cursor: pointer;
}

#excelTable {
    margin-top: 20px;
}

footer {
    background-color: #333;
    color: white;
    text-align: center;
    padding: 20px 0;
}

footer .social-links {
    list-style: none;
    padding: 0;
    margin: 0;
    display: flex;
    justify-content: center;
    gap: 20px;
}

footer .social-links li a {
    color: white;
    text-decoration: none;
    font-size: 1.2em;
    transition: color 0.3s;
}

footer .social-links li a:hover {
    color: #ff6f61;
}

/* Handsontable custom styles */
.handsontable .htCore td, .handsontable .htCore th {
    white-space: normal !important; /* Enable text wrapping */
}