import flet as ft
import sqlite3
import docxtpl
import os
import num2words
import requests


# 250.10 | 250,10 -> Двести пятьдесят рублей, десять копеек
def num_to_text(num: str):
    num = str(num)
    num = num.strip().replace(',', '.')
    num = float(num)
    num_to_word = num2words.num2words(num, lang='ru', to='currency', currency='RUB').capitalize()
    return num_to_word

# Если flag_field == True, то делает return словаря с переменными
def save_file_docx(categories: [], file_docx: docxtpl.DocxTemplate, path: str, flag_field: bool):
    fields = {}
    name_file_from_user = 'filled_out.docx'
    ban_button = [
        ("Название файла", "Название файла"),
        ("Сумма прописью", "Сумма прописью"),
        ("Склонить слова", "Склонить слова")
    ]
    for category in categories:
        # Проходимся по всем TextField (пропускаем первый, потому что там Text - название категории)
        for value_var in category.controls[1:]:
            var_in_file, var_after_replace = value_var.data
            # Проверка, чтобы пропустить поле с названием файла
            if value_var.data not in ban_button:
                if value_var.value:
                    fields[var_in_file] = value_var.value
                else:
                    fields[var_in_file] = "{{ " + var_in_file + " }}"
            elif value_var.data == ("Название файла", "Название файла"):
                if value_var.value:
                    name_file_from_user = f"{value_var.value}.docx"
    if flag_field == False:
        file_docx.render(fields)

        # Получаем директорию без файла, только последней папки
        directory = os.path.dirname(path)
        # Формируем полный путь выходного файла
        output_path = os.path.join(directory, name_file_from_user)

        file_docx.save(output_path)
    else:
        return fields


# Возвращает форму строки в родительном падеже через API Morpher.
def get_genitive(text: str) -> str | None:
    """
    Возвращает форму строки в родительном падеже через API Morpher.

    :param text: Исходное слово или ФИО.
    :return: Строка в родительном падеже или None, если не удалось получить.
    """
    url = "https://ws3.morpher.ru/russian/declension"
    params = {
        "s": text,
        "format": "json"
    }
    headers = {
        "User-Agent": "DeclensionAPIClient"
    }

    response = requests.get(url, params=params, headers=headers)
    response.raise_for_status()  # выбросит HTTPError при ошибке
    data = response.json()

    # 'Р' — родительный падеж
    return data.get("Р")


def main(page: ft.Page):
    page.title = "DocumentFill"
    page.theme_mode = "dark"
    page.scroll = "auto"
    # page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.window.height = 350 * 2
    page.window.width = 400 * 2


    flag_file_selected = False

    # Сохраняет название шаблона в переменную selected_files
    def pick_files_result(e: ft.FilePickerResultEvent):
        try:
            if e.files:
                nonlocal flag_file_selected
                flag_file_selected = True
                # e.files структура: [FilePickerFile(name='test.docx', path='C:\\Users', size=17283, id=0)]
                selected_files.value = e.files[0].name
                path = e.files[0].path

                # Получение переменных из файла docx
                file_docx = docxtpl.DocxTemplate(path)
                variables = sorted(file_docx.get_undeclared_template_variables())

                confirm_save_var = ft.AlertDialog(
                    modal=True,
                    title=ft.Text("Подтверждение"),
                    content=ft.Text("Заполнены не все поля. Все равно сохранить?"),
                    actions=[
                        ft.TextButton("Да", on_click=lambda e: confirm_save(confirm_save_var)),
                        ft.TextButton("Нет", on_click=lambda e: close_save(confirm_save_var)),
                    ],
                    actions_alignment=ft.MainAxisAlignment.END
                )

                def confirm_save(dialog):
                    page.close(dialog)
                    nonlocal categories

                    try:
                        save_file_docx(categories, file_docx, path, False)
                        page.open(ft.SnackBar(ft.Text("Файл сохранен")))
                        page.update()
                    except Exception as e:
                        page.open(ft.SnackBar(ft.Text("Произошла ошибка")))
                        page.update()

                def close_save(dialog):
                    page.close(dialog)

                ban_button = [
                    ("Название файла", "Название файла"),
                    ("Сумма прописью", "Сумма прописью"),
                    ("Склонить слова", "Склонить слова")
                ]

                def on_save(e):
                    nonlocal categories
                    cnt_var = 0
                    cnt_var_filled = 0

                    nonlocal ban_button

                    for category in categories:
                        # Проходимся по всем TextField (пропускаем первый, потому что там Text - название категории)
                        for value_var in category.controls[1:]:
                            var_in_file, var_after_replace = value_var.data
                            # Проверка, чтобы пропустить поле с названием файла
                            if value_var.data not in ban_button:
                                cnt_var += 1
                                if value_var.value:
                                    cnt_var_filled += 1

                    if cnt_var == cnt_var_filled:
                        try:
                            save_file_docx(categories, file_docx, path, False)
                            page.open(ft.SnackBar(ft.Text("Файл сохранен")))
                            page.update()
                        except Exception as e:
                            page.open(ft.SnackBar(ft.Text("Произошла ошибка")))
                            page.update()
                    else:
                        page.open(confirm_save_var)

                """
                Категории
                """

                organization = ft.Column([ft.Text('Заказчик')])
                manager_info = ft.Column([ft.Text('Руководитель')])
                company_details = ft.Column([ft.Text('Реквизиты')])
                object_info = ft.Column([ft.Text('Объект')])
                contract_info = ft.Column([ft.Text('Информация о договоре')])
                other = ft.Column([ft.Text('Общий раздел')])

                name_file_out = ft.Column([ft.Text('Название файла')])
                btn_tf_name_file = ft.TextField(label="Название файла")
                btn_tf_name_file.data = ("Название файла", "Название файла")
                name_file_out.controls.append(btn_tf_name_file)

                btn_save = ft.ElevatedButton("Сохранить", on_click=on_save, width=150)

                btn_num_to_word = ft.ElevatedButton("Сумма прописью", width=150, on_click=lambda e: num_to_text(e))
                btn_num_to_word.data = ("Сумма прописью", "Сумма прописью")

                btn_case = ft.ElevatedButton("Склонить слова", width=150)
                btn_case.data = ("Склонить слова", "Склонить слова")

                row_buttons = ft.Row(
                    controls=[
                        btn_num_to_word,
                        btn_case,
                    ],
                    wrap=True,
                )

                """
                Теги в файле docx
                ==================================
                organization_
                manager_
                company_
                object_
                contract_
                sum_in_words_
                word_gent_
                ==================================
                """

                def num_to_text(e):
                    nonlocal ban_button
                    fields = save_file_docx(categories, file_docx, path, True)
                    var_list = [var for var in fields.keys()]
                    # print(var_list)
                    for category in categories:
                        # Проходимся по всем TextField (пропускаем первый, потому что там Text - название категории)
                        for value_var, in category.controls[1:]:
                            var_in_file, var_after_replace = value_var.data
                            # Проверка, чтобы пропустить поле с названием файла
                            if value_var.data not in ban_button:
                                ...
                                # print(value_var)



                def check_sum_in_words__(text: str):
                    if text.startswith('sum_in_words_'):
                        text = text.replace('sum_in_words_', '')
                    elif text.startswith('word_gent_'):
                        text = text.replace('word_gent_', '')
                    return text

                def add_var_in_category():
                    for variable in variables:
                        new_variable = variable.strip()
                        if variable.startswith('organization_'):
                            new_variable = variable.replace('organization_', '')
                            new_variable = check_sum_in_words__(new_variable)
                            new_variable = new_variable.replace('_', ' ')
                            tf = ft.TextField(label=new_variable, multiline=True, )
                            tf.data = (variable, new_variable)
                            organization.controls.append(tf)

                        elif variable.startswith('manager_'):
                            new_variable = variable.replace('manager_', '')
                            new_variable = check_sum_in_words__(new_variable)
                            new_variable = new_variable.replace('_', ' ')
                            tf = ft.TextField(label=new_variable, multiline=True, )
                            tf.data = (variable, new_variable)
                            manager_info.controls.append(tf)

                        elif variable.startswith('company_'):
                            new_variable = variable.replace('company_', '')
                            new_variable = check_sum_in_words__(new_variable)
                            new_variable = new_variable.replace('_', ' ')
                            tf = ft.TextField(label=new_variable, multiline=True, )
                            tf.data = (variable, new_variable)
                            company_details.controls.append(tf)

                        elif variable.startswith('object_'):
                            new_variable = variable.replace('object_', '')
                            new_variable = check_sum_in_words__(new_variable)
                            new_variable = new_variable.replace('_', ' ')
                            tf = ft.TextField(label=new_variable, multiline=True, )
                            tf.data = (variable, new_variable)
                            object_info.controls.append(tf)

                        elif variable.startswith('contract_'):
                            new_variable = variable.replace('contract_', '')
                            new_variable = check_sum_in_words__(new_variable)
                            new_variable = new_variable.replace('_', ' ')
                            tf = ft.TextField(label=new_variable, multiline=True, )
                            tf.data = (variable, new_variable)
                            contract_info.controls.append(tf)

                        else:
                            new_variable = check_sum_in_words__(new_variable)
                            new_variable = new_variable.replace('_', ' ')
                            tf = ft.TextField(label=new_variable, multiline=True, )
                            tf.data = (variable, new_variable)
                            other.controls.append(tf)

                add_var_in_category()

                categories = [
                    organization,
                    manager_info,
                    company_details,
                    object_info,
                    contract_info,
                    other,
                    name_file_out,
                    row_buttons
                ]

                # Добавление информации на страницу
                def add_info():
                    # Удаляем все кроме первых двух элементов
                    page.controls = page.controls[:2]

                    page.controls.append(ft.Divider(10))

                    # Если в блоках содержится не только один ft.Text() с названием блока, то добавляем на страницу
                    if len(organization.controls) > 1:
                        page.controls.append(organization)
                        page.controls.append(ft.Divider(10))

                    if len(manager_info.controls) > 1:
                        page.controls.append(manager_info)
                        page.controls.append(ft.Divider(10))

                    if len(company_details.controls) > 1:
                        page.controls.append(company_details)
                        page.controls.append(ft.Divider(10))

                    if len(object_info.controls) > 1:
                        page.controls.append(object_info)
                        page.controls.append(ft.Divider(10))

                    if len(contract_info.controls) > 1:
                        page.controls.append(contract_info)
                        page.controls.append(ft.Divider(10))

                    if len(other.controls) > 1:
                        page.controls.append(other)
                        page.controls.append(ft.Divider(10))

                    page.controls.append(name_file_out)

                    page.controls.append(ft.Divider(10))
                    # Временно отключено
                    # page.controls.append(row_buttons)
                    page.controls.append(btn_save)

                add_info()
                page.update()

            else:
                page.open(ft.SnackBar(ft.Text("Шаблон не выбран!")))
                page.update()
            selected_files.update()
            page.update()
        except Exception as e:
            page.open(ft.SnackBar(ft.Text("Произошла ошибка. Скорее всего в названии переменных есть пробелы")))
            page.update()

    pick_files_dialog = ft.FilePicker(on_result=pick_files_result)
    # selected_files содержит имя загруженного файла
    selected_files = ft.Text()

    page.overlay.append(pick_files_dialog)

    # Диалог с подтверждением
    confirm_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("Подтверждение"),
        content=ft.Text("Файл уже загружен. Заменить?"),
        actions=[
            ft.TextButton("Да", on_click=lambda e: confirm_replace(confirm_dialog)),
            ft.TextButton("Нет", on_click=lambda e: close_confirm(confirm_dialog))
        ],
        actions_alignment=ft.MainAxisAlignment.END
    )

    # Закрыть диалог
    def close_confirm(dialog):
        page.close(dialog)

    # Выполнить замену файла
    def confirm_replace(dialog):
        page.close(dialog)
        pick_files_dialog.pick_files(
            allow_multiple=False,
            allowed_extensions=["docx"]
        )

    # Выбор файла
    def choice_file_click(e):
        nonlocal flag_file_selected
        # Если файл уже загружен - спрашиваем пользователя о замене файла
        if flag_file_selected:
            page.open(confirm_dialog)
        # Иначе просто открываем файл
        else:
            pick_files_dialog.pick_files(
                allow_multiple=False,
                allowed_extensions=["docx"]
            )

    """
    Кнопки (btn)
    """

    btn_choice_file = ft.ElevatedButton(
        "Выбрать шаблон",
        icon=ft.Icons.UPLOAD_FILE,
        on_click=lambda e: choice_file_click(e)
        # on_click=lambda _: pick_files_dialog.pick_files(
        #     allow_multiple=False,
        #     allowed_extensions=["docx"]
        # )
    )

    """
    Контент на странице
    """
    content = ft.Row(
        controls=[
            btn_choice_file,
            selected_files,
        ],
        wrap=True,
    )

    def main_page():
        appbar = ft.AppBar(
            title=ft.Text("Главная страница")
        )
        return appbar

    """
    Добавление на страницу
    """

    page.controls.append(content)
    page.controls.append(main_page())
    page.update()


ft.app(target=main)
