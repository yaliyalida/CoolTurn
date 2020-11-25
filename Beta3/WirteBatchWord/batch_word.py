import batch_word_text
import batch_word_table


if __name__ == "__main__":
    print("请选择:\n1、批量写入文本形式\n2、批量写入表格形式")
    optional = ('1','2')
    while True:
        try:
            choice = input()
            if choice not in optional:
                raise ValueError('请按照提示输入')
            break
        except Exception as err:
            print(err)

    if choice == '1':  # 批量写入文本形式
        formatpath = 'template_text.docx'
        batch_word_text.main(formatpath)
    elif choice == '2':  # 批量写入表格形式
        formatpath = 'template_table.docx'
        batch_word_table.main(formatpath)

