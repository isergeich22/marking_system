const home_button = document.querySelector('#home')
const import_button = document.querySelector('#import')
const cis_actions_button = document.querySelector('#cis_actions')
const other_actions_button = document.querySelector('#other_actions')
const import_control = document.querySelector('.import-control')
const marking_control = document.querySelector('.marking-control')
const other_control = document.querySelector('.other-control')
const nav_items = document.querySelectorAll('.nav-item')

home_button.addEventListener('click', () => {
    [import_control.style.display, marking_control.style.display, other_control.style.display] = ['none', 'none', 'none']
    nav_items.forEach(el => {
        el.classList.remove('active')
    })
    home_button.classList.add('active')
})

import_button.addEventListener('click', () => {
    import_control.style.display = 'flex'
    marking_control.style.display = 'none'
    other_control.style.display = 'none'
    nav_items.forEach(el => {
        el.classList.remove('active')
    })
    import_button.classList.add('active')
})

cis_actions_button.addEventListener('click', () => {
    marking_control.style.display = 'flex'
    import_control.style.display = 'none'
    other_control.style.display = 'none'
    nav_items.forEach(el => {
        el.classList.remove('active')
    })
    cis_actions_button.classList.add('active')
})

other_actions_button.addEventListener('click', () => {

    other_control.style.display = 'flex'
    marking_control.style.display = 'none'
    import_control.style.display = 'none'
    nav_items.forEach(el => {
        el.classList.remove('active')
    })
    other_actions_button.classList.add('active')

})

const table_header = document.querySelector('.marks-table-header')
const buttonTop = document.querySelector('#top')
const multipleList = document.querySelector('.multiple-list')
const statusList = document.querySelector('.status-list')
const statusRows = document.querySelectorAll('#status')

window.addEventListener('scroll', () => {

    if(scrollY > 100) {
        table_header.classList.add('--pinned')
        buttonTop.style.display = 'block'
    } else {
        table_header.classList.remove('--pinned')
        buttonTop.style.display = 'none'
    }

    buttonTop.addEventListener('click', () => {
        document.documentElement.scrollTop = 0
    })

})

if(window.location.href.indexOf('home') >= 0) {
    multipleList.addEventListener('click', () => {

        const css = window.getComputedStyle(statusList)
        if(statusList.style.display == 'none' ||  css.display == 'none') {
            statusList.style.display = 'block'
        } else {
            statusList.style.display = 'none'
        }
        

    })
}

statusRows.forEach(el => {
    if(el.innerHTML == 'В обороте') {
        el.style.color = '#36AD60'
    } else if(el.innerHTML == 'Нанесен') {
        el.style.color = 'rgb(240, 141, 27)'
    } else if(el.innerHTML == 'Выбыл') {
        el.style.color = 'rgb(122, 129, 155)'
    }
})

//pagination logic begin {
if(window.location.href.indexOf('filter') < 0 && window.location.href.indexOf('yandex') < 0) {
    
    const beginButton = document.querySelector('#begin')
    const prevButton = document.querySelector('#prev')
    const nextButton = document.querySelector('#next')
    const lastButton = document.querySelector('#last')

    const orderInput = document.querySelectorAll('#order')
    const submitButton = document.querySelectorAll('.order-number')
    const marksSpan = document.querySelectorAll('#mark')

    const [prevIcon, nextIcon] = [document.querySelector('#prev-icon'), document.querySelector('#next-icon')]

    let href = lastButton.getAttribute('href')
    let lastPage = parseInt(href.split('=').pop())

    if(parseInt(window.location.href.split('=').pop()) == 1 || window.location.href.charAt(window.location.href.length - 1) == 'e' || window.location.href.charAt(window.location.href.length - 1) == 'D') {
        
        beginButton.removeAttribute('href')
        prevButton.removeAttribute('href')
        beginButton.classList.add('disabled')
        prevButton.classList.add('disabled')
        prevIcon.style.fill = '#c4c6c9'
        prevIcon.style.cursor = 'text'
        if(window.location.href.indexOf('APPLIED') >= 0) {
            nextButton.setAttribute('href', `http://localhost:3030/home/APPLIED?page=2`)
        }
        if(window.location.href.indexOf('RETIRED') >= 0) {
            nextButton.setAttribute('href', `http://localhost:3030/home/RETIRED?page=2`)
        }
        if(window.location.href.indexOf('INTRODUCED') >= 0) {
            nextButton.setAttribute('href', `http://localhost:3030/home/INTRODUCED?page=2`)
        } 
        if(window.location.href.indexOf('APPLIED') < 0 && window.location.href.indexOf('RETIRED') < 0 && window.location.href.indexOf('INTRODUCED') < 0){
            nextButton.setAttribute('href', `http://localhost:3030/home?page=2`)
        }

            
    } else {

        let pageNumber = parseInt(window.location.href.split('=').pop())

        prevButton.classList.remove('disabled')
        prevIcon.style.fill = '#63666a'
        prevIcon.style.cursor = 'pointer'

        if(window.location.href.indexOf('APPLIED') >= 0) {
            prevButton.setAttribute('href', `http://localhost:3030/home/APPLIED?page=${pageNumber - 1}`)
            nextButton.setAttribute('href', `http://localhost:3030/home/APPLIED?page=${pageNumber + 1}`)
        }
        if(window.location.href.indexOf('INTRODUCED') >= 0) {
            prevButton.setAttribute('href', `http://localhost:3030/home/INTRODUCED?page=${pageNumber - 1}`)
            nextButton.setAttribute('href', `http://localhost:3030/home/INTRODUCED?page=${pageNumber + 1}`)
        }
        if(window.location.href.indexOf('RETIRED') >= 0) {
            prevButton.setAttribute('href', `http://localhost:3030/home/RETIRED?page=${pageNumber - 1}`)
            nextButton.setAttribute('href', `http://localhost:3030/home/RETIRED?page=${pageNumber + 1}`)
        }
        if(window.location.href.indexOf('APPLIED') < 0 && window.location.href.indexOf('RETIRED') < 0 && window.location.href.indexOf('INTRODUCED') < 0){
            prevButton.setAttribute('href', `http://localhost:3030/home?page=${pageNumber - 1}`)
            nextButton.setAttribute('href', `http://localhost:3030/home?page=${pageNumber + 1}`)
        }

    }

    if(lastPage == parseInt(window.location.href.split('=').pop())) {

        lastButton.removeAttribute('href')
        lastButton.classList.add('disabled')
        nextButton.removeAttribute('href')
        nextButton.classList.add('disabled')
        nextIcon.style.fill = '#c4c6c9'
        nextIcon.style.cursor = 'text'
        if(window.location.href.indexOf('APPLIED') >= 0) {
            prevButton.setAttribute('href', `http://localhost:3030/home/APPLIED?page=${lastPage - 1}`)
        }
        if(window.location.href.indexOf('INTRODUCED') >= 0) {
            prevButton.setAttribute('href', `http://localhost:3030/home/INTRODUCED?page=${lastPage - 1}`)
        }
        if(window.location.href.indexOf('RETIRED') >= 0) {
            prevButton.setAttribute('href', `http://localhost:3030/home/RETIRED?page=${lastPage - 1}`)
        }
        if(window.location.href.indexOf('APPLIED') < 0 && window.location.href.indexOf('RETIRED') < 0 && window.location.href.indexOf('INTRODUCED') < 0){
            prevButton.setAttribute('href', `http://localhost:3030/home?page=${lastPage - 1}`)
        }

    }

}
// } pagination logic end 

//status-filter logic begin {
const multipleStatus = document.querySelector('.multiple-status')
const multipleItems = document.querySelectorAll('.list-item')
const showButton = document.querySelector('.show-button')
const showAnchor = document.querySelector('#show-anchor')

if(window.location.href.indexOf('home') >= 0) {
    multipleItems.forEach(el => {

        el.addEventListener('click', () => {

            multipleStatus.innerHTML = el.innerHTML
            multipleStatus.style.color = '#181F3E;'
            showButton.style.display = 'inline-block'
            if(multipleStatus.innerHTML == 'Нанесен') {

                showAnchor.setAttribute('href', `http://localhost:3030/home/APPLIED`)

            }

            if(multipleStatus.innerHTML == 'В обороте') {

                showAnchor.setAttribute('href', `http://localhost:3030/home/INTRODUCED`)

            }

            if(multipleStatus.innerHTML == 'Выбыл') {

                showAnchor.setAttribute('href', `http://localhost:3030/home/RETIRED`)

            }
            

        })

    })
}
// } filter logic end

//kiz-and-gtin-filter logic begin {

if(window.location.href.indexOf('yandex') < 0) {
    
    const searchButton = document.querySelector('#search')
    const searchField = document.querySelector('.search-input')

    searchButton.addEventListener('click', () => {

        showButton.style.display = 'inline-block'

        let str = searchField.value

        if(str.indexOf('<') >= 0) {
            str = str.replace(/</g, '&lt;')
        } else if(str.indexOf('&') >= 0) {
            str = str.replace(/&/g, '&amp;')
        }
        
        if(searchField.value.length == 31) {   
            showAnchor.setAttribute('href', `http://localhost:3030/filter?cis=${str}`)
        }

        if(searchField.value.length == 14) {
            showAnchor.setAttribute('href', `http://localhost:3030/filter?gtin=${searchField.value}`)
        }

    })

}

// } kiz-and-gtin-filter logic end

// yandex-logic begin {

    if(window.location.href.indexOf('yandex') >= 0) {
        const ref = document.querySelector('.input-form__ref')
        console.log(ref)
        const input = document.querySelector('.input-form__input')
        console.log(input)

        input.addEventListener('input', () => {

                let href = `/yandex?cis=${input.value.replace(/&/g, 'AND')}`

                href = href.replace(/%/g, 'percent')

                ref.setAttribute('href', href)

        })
    }

// } yandex-logic end