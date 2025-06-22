"""Модуль для парсинга задач, глав и информации об авторе из документа Word.

Этот модуль предоставляет функциональность для обработки
документов Word (.docx) и извлечения:
- Задач с ответами
- Иерархии глав
- Информации об авторе

Извлеченные данные экспортируются в файл Excel с отдельными
листами для каждого типа данных.
"""

import re

import pandas
import docx

from constants import (
    RE_ANSWER_TEXT,
    RE_CHAPTER,
    RE_PART_ANSWER,
    RE_SUPREME_CHAPTER,
    RE_TASK,
    RE_TASK_ID,
    ABSENT
)


def find_next_task(data, tasks_data, next_index, previous_id, paragraph):
    """Рекурсивно находит следующую задачу в документе.

    Args:
        data (list): Список текстовых параграфов из документа.
        tasks_data (list): Список для хранения извлеченных данных о задачах.
        next_index (int): Текущий индекс в списке данных для начала поиска.
        previous_id (str): ID предыдущей задачи для ссылки.
        paragraph (int): ID текущей главы.

    Returns:
        int: Индекс следующего элемента после последней найденной задачи.
    """
    text = data[next_index]
    task_match = re.search(RE_TASK, text)
    if task_match:
        if re.search(RE_SUPREME_CHAPTER, text):
            return next_index
        tasks_data.append(get_task_data(task_match, paragraph, previous_id))
        next_index = find_next_task(data, tasks_data,
                                    next_index + 1, previous_id, paragraph)
    return next_index


def find_next_chapter(data, next_index):
    """Находит следующую главу в документе.

    Args:
        data (list): Список текстовых параграфов из документа.
        next_index (int): Текущий индекс в списке данных для начала поиска.

    Returns:
        tuple: (Match-объект с найденной главой, индекс текущего элемента)
    """
    text = data[next_index]
    chapter_match = chapter_match = re.search(RE_CHAPTER, text)
    return chapter_match, next_index


def get_chapter_data(id, match, parent=0):
    """Формирует данные о главе.

    Args:
        id (int): Уникальный идентификатор главы.
        match (Match): Match-объект с найденной главой.
        parent (int, optional): ID родительской главы. По умолчанию 0.

    Returns:
        dict: Словарь с данными о главе.
    """
    if not parent:
        return {
            'id': id,
            'name': f'{match.group(1)}{match.group(2)}',
            'parent': parent
        }
    else:
        return {
            'id': id,
            'name': f'{match.group(1)}{match.group(2)}{match.group(3)}',
            'parent': parent
        }


def get_task_data(match, paragraph, previous=None,
                  exclusive=False, classes='5;6', level=1, topic_id=1):
    """Формирует данные о задаче.

    Args:
        match (Match): Match-объект с найденной задачей.
        paragraph (int): ID текущей главы.
        previous (str, optional): ID предыдущей задачи. По умолчанию None.
        exclusive (bool, optional): Флаг исключительной
        задачи. По умолчанию False.
        classes (str, optional): Классы, для которых предназначена
        задача. По умолчанию '5;6'.
        level (int, optional): Уровень сложности. По умолчанию 1.
        topic_id (int, optional): ID темы. По умолчанию 1.

    Returns:
        dict: Словарь с данными о задаче.
    """
    data = {
        'classes': classes,
        'paragraph': paragraph,
        'topic_id': topic_id,
        'level': level
    }
    if exclusive:
        data['id_tasks_book'] = (match.group(1))[:-1]
        data['task'] = match.group(2)
    elif match.group(1) is None:
        if not previous.endswith('.'):
            previous = previous + '.'
        data['id_tasks_book'] = (previous + match.group(4))[:-1]
        data['task'] = match.group(5)
    else:
        data['id_tasks_book'] = (match.group(1) + match.group(2))[:-1]
        data['task'] = match.group(3)

    return data


def get_answer_text(task_id, text):
    """Извлекает текст ответа для задачи.

    Args:
        task_id (str): ID задачи.
        text (str): Текст для поиска ответа.

    Returns:
        str: Текст ответа или ABSENT, если ответ не найден.
    """
    pattern = fr'{task_id}' + RE_ANSWER_TEXT
    answer_match = re.search(pattern, text)
    if answer_match:
        return answer_match.group(1).strip()
    return ABSENT


def get_answer_part(task_part_id, text):
    """Извлекает часть ответа для подзадачи.

    Args:
        task_part_id (str): ID подзадачи.
        text (str): Текст для поиска ответа.

    Returns:
        str: Текст части ответа или ABSENT, если ответ не найден.
    """
    pattern = fr'{task_part_id}' + RE_PART_ANSWER
    answer_part_match = re.search(pattern, text)
    if answer_part_match:
        return answer_part_match.group(1)[:answer_part_match.end()].strip()
    return ABSENT


def answer_parser(data, tasks_data):
    """Парсит ответы для задач.

    Args:
        data (list): Список текстовых параграфов с ответами.
        tasks_data (list): Список данных о задачах.

    Returns:
        list: Обновленный список задач с добавленными ответами.
    """
    index = 0
    answers_text = str.join('\n', data)
    current_task_num = 0
    while index < len(tasks_data[:]):
        task = tasks_data[index]
        task_id_match = re.search(RE_TASK_ID, task['id_tasks_book'])
        if task_id_match.group(1):
            if int(task_id_match.group(1)) >= current_task_num:
                current_task_num = int(task_id_match.group(1))
                answer = get_answer_text(current_task_num, answers_text)
                if answer != ABSENT:
                    answer_part = get_answer_part(task_id_match.group(2),
                                                  answer)
                    tasks_data[index]['answer'] = answer_part
                else:
                    tasks_data[index]['answer'] = answer

        else:
            current_task_num = int(task_id_match.group())
            next_task = tasks_data[index + 1]
            next_task_id_match = re.search(RE_TASK_ID,
                                           next_task['id_tasks_book'])
            if next_task_id_match.group(1) is None:
                answer = get_answer_text(current_task_num, answers_text)
                tasks_data[index]['answer'] = answer
                current_task_num = int(next_task_id_match.group())
                next_answer = get_answer_text(current_task_num, answers_text)
                tasks_data[index + 1]['answer'] = next_answer
                index += 1
            elif int(next_task_id_match.group(1)) > current_task_num:
                answer = get_answer_text(current_task_num, answers_text)
                tasks_data[index]['answer'] = answer
                current_task_num = int(next_task_id_match.group(1))
                next_answer = get_answer_text(current_task_num, answers_text)
                if next_answer != ABSENT:
                    next_answer_part = get_answer_part(
                        next_task_id_match.group(2),
                        next_answer)
                    tasks_data[index + 1]['answer'] = next_answer_part
                else:
                    tasks_data[index + 1]['answer'] = next_answer
                index += 1
            elif int(next_task_id_match.group(1)) == current_task_num:
                tasks_data[index]['answer'] = ABSENT
                next_answer = get_answer_text(current_task_num, answers_text)
                if next_answer != ABSENT:
                    next_answer_part = get_answer_part(
                        next_task_id_match.group(2), next_answer)
                    tasks_data[index + 1]['answer'] = next_answer_part
                index + 1
        index += 1
    return tasks_data


def parser(doc):
    """Основная функция парсинга документа.

    Args:
        doc (Document): Объект документа Word.

    Returns:
        tuple: Кортеж с тремя списками:
            - Данные о задачах
            - Данные о главах
            - Данные об авторе
    """
    chapters_data = []
    tasks_data = []
    answers_text = []
    authors_data = [{
        'name': 'Текстовые задачи по математике. 5–6 классы / ' +
        'А. В. Шевкин. — 3-е изд., перераб. — М. : '
        'Илекса, 2024. — 160 с. : ил.',
        'description': 'Сборник включает текстовые задачи по разделам '
        'школьной математики: натуральные числа, дроби, пропорции, '
        'проценты, уравнения. Ко многим задачам даны ответы или советы с '
        'чего начать решения. Решения некоторых задач приведены в качестве '
        'образцов в основном тексте книги или в разделе '
        '«Ответы, советы, решения». '
        'Материалы сборника можно использовать как '
        'дополнение к любому действующему '
        'учебнику. При подготовке этого издания добавлены новые '
        'задачи и решения некоторых '
        'задач. Пособие предназначено для учащихся 5–6 классов '
        'общеобразовательных школ, учителей, '
        'студентов педагогических вузов.',
        'topic_id': 1,
        'classes': '5;6'
    }]

    data = []
    index = 0
    current_chapter = 0
    supreme_chapter_id = 0
    current_id = 1
    for item in doc.iter_inner_content():
        data.append(item.text.strip())
    while index < len(data):
        text = data[index]
        supreme_chapter_match = re.search(RE_SUPREME_CHAPTER, text)
        if supreme_chapter_match:
            its_next_chapter = int(
                supreme_chapter_match.group(1)[:-1]) == current_chapter + 1
            task_match = re.search(RE_TASK, text)
            if its_next_chapter and task_match is None:
                chapter_match, chapter_index = find_next_chapter(
                    data, index + 1)
                if chapter_match:
                    chapters_data.append(get_chapter_data(
                        current_id,
                        supreme_chapter_match))
                    supreme_chapter_id = current_id
                    current_id += 1

                    chapters_data.append(get_chapter_data(current_id,
                                                          chapter_match,
                                                          supreme_chapter_id))
                    current_id += 1

                    current_chapter = int(supreme_chapter_match.group(1)[:-1])
                    index = chapter_index + 1
                else:
                    chapters_data.append(get_chapter_data(
                        current_id,
                        supreme_chapter_match))
                    supreme_chapter_id = current_id
                    current_id += 1

                    current_chapter = int(supreme_chapter_match.group(1)[:-1])
                    index += 1
                continue
            else:
                chapter_match = re.search(RE_CHAPTER, text)
                if chapter_match:
                    chapters_data.append(get_chapter_data(current_id,
                                                          chapter_match,
                                                          supreme_chapter_id))
                    current_id += 1

                    index += 1
                    continue
                paragraph = chapters_data[-1]['id']
                if task_match:
                    previous_id = (
                        tasks_data[-1]['id_tasks_book'] if (
                            task_match.group(4)) else task_match.group(1))
                    tasks_data.append(get_task_data(
                        task_match, paragraph, previous_id))
                    index = find_next_task(data, tasks_data, index + 1,
                                           task_match.group(1),
                                           paragraph=paragraph)
                else:
                    tasks_data.append(get_task_data(supreme_chapter_match,
                                                    paragraph,
                                                    exclusive=True))
                    index += 1
            continue
        chapter_match = re.search(RE_CHAPTER, text)
        if chapter_match:
            chapters_data.append(get_chapter_data(current_id,
                                                  chapter_match,
                                                  supreme_chapter_id))
            current_id += 1

            index += 1
            continue
        task_match = re.search(RE_TASK, text)
        if task_match:
            previous_id = (
                tasks_data[-1]['id_tasks_book'] if (
                    task_match.group(4)) else task_match.group(1))
            paragraph = chapters_data[-1]['id']
            tasks_data.append(get_task_data(
                task_match, paragraph, previous_id))
            index = find_next_task(
                data, tasks_data, index + 1,  previous_id, paragraph)
        elif text == 'Ответы и советы':
            index += 1
            while index < len(data[:]):
                text = data[index]
                if text.lower().strip() == 'оглавление':
                    break
                answers_text.append(text)
                index += 1
            break
        else:
            raise ValueError(f'index:{index}    {text}')

    if answers_text != []:
        tasks_data = answer_parser(answers_text, tasks_data)
    return tasks_data, chapters_data, authors_data


if __name__ == '__main__':
    docx_path = input('Укажите путь к .docx-файду:\n')
    excel_path = input('Введите название для excel файла:\n')

    document = docx.Document(docx_path)
    tasks_data, chapters_data, authors_data = parser(document)

    tasks_columns_order = [
        'id_tasks_book', 'task', 'answer', 'classes',
        'paragraph', 'topic_id', 'level'
    ]
    author_columns_order = ['author', 'description', 'topic_id', 'classes']

    chapters_columns_order = ['id', 'name', 'parent']

    tasks_df = pandas.DataFrame(tasks_data, columns=tasks_columns_order)
    authors_df = pandas.DataFrame(authors_data, columns=author_columns_order)
    chapters_df = pandas.DataFrame(
        chapters_data, columns=chapters_columns_order)
    with pandas.ExcelWriter(excel_path, engine='openpyxl') as writer:
        tasks_df.to_excel(writer, sheet_name='tasks', index=False)
        authors_df.to_excel(writer, sheet_name='author', index=False)
        chapters_df.to_excel(
            writer, sheet_name='table_of_contents', index=False)
