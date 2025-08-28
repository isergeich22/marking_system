const headerComponent = `<!DOCTYPE html>
                            <html lang="en">
                            <head>
                                <meta charset="UTF-8">
                                <meta http-equiv="X-UA-Compatible" content="IE=edge">
                                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                                <link rel="stylesheet" href="/css/styles.css" type="text/css">
                                <link rel="shortcut icon" type="image/png" href="/favicon.png">
                                <link
                                    rel="stylesheet"
                                    href="https://cdn.jsdelivr.net/npm/bulma@1.0.4/css/bulma.min.css"
                                >`

const navComponent = `<header class="header">
                        <nav>
                            <img src="/img/chestnyj_znak.png" alt="честный знак">
                            <p class="nav-item" id="home"><a href="http://localhost:3030/home">Главная</a></p>
                            <p class="nav-item" id="import">Создание импорт-файлов</p>
                            <p class="nav-item" id="cis_actions">Действия с КИЗ</p>
                            <p class="nav-item" id="other_actions">Другие действия</p>
                        </nav>                    
                    </header>`

const footerComponent = `   <button id="top" class="button-top">
                            <svg width="24" height="24" fill="none" xmlns="http://www.w3.org/2000/svg">
                                <g clip-path="url(#ArrowLongUp_large_svg__clip0_35331_5070)">
                                    <path d="M12 2v20m0-20l7 6.364M12 2L5 8.364" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"></path>
                                </g><defs><clipPath id="ArrowLongUp_large_svg__clip0_35331_5070"><path fill="#fff" transform="rotate(90 12 12)" d="M0 0h24v24H0z">
                                </path></clipPath></defs></svg>
                            </button>    
                            <script src="/script.js"></script>
                            </body>
                        </html>`

module.exports = {

    headerComponent,
    navComponent,
    footerComponent

}