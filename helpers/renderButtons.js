function renderImportButtons(array) {

    let str = ''

    for(let i = 0; i < array.length; i++) {

        if(array[i] === 'stocks') {

            str += `<button class="button-import">
                        <a href="/stocks" target="_blank">Создать импорт для остатков</a>
                    </button>`

        }

        if(array[i] === 'wb') {

            str += `<button class="button-import">
                        <a href="/wildberries" target="_blank">Создать импорт для ${array[i]}</a>
                    </button>`

        }

        if(array[i] !== 'wb' && array[i] !== 'stocks') {
            str += `<button class="button-import">
                        <a href="/${array[i]}" target="_blank">Создать импорт для ${array[i]}</a>
                    </button>`
        }

    }

    str += `   </div>`

    return str

}

function renderMarkingButtons() {
    return `<div class="marking-control">
                <button class="marking-button remarking-button"><a href="/input_own" target="_blank">Ввод в оборот (Производство РФ)</a></button>
                <button class="marking-button distance-button"><a href="/sale_ozon" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                <button class="marking-button distance-button"><a href="/sale_wb" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                <button class="marking-button distance-button"><a href="/wildberries/set_marks" target="_blank">Подстановка маркировки (Wildberries)</a></button>
            </div>`
}

function renderExtraButtons() {

    return `<div class="other-control">
                <button class="other-button mark-stocks"><a href="/test_features" target="_blank">Обновить краткий отчет</a></button>
                <button class="other-button mark-stocks"><a href="/personal_orders" target="_blank">Создать персональный заказ</a></button>
            </div>`

}

module.exports = { renderImportButtons, renderMarkingButtons, renderExtraButtons }
