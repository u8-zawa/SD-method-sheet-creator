from random import shuffle

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

from settings import WORD_PASS, NEW_WORD_PASS

adjective_pair_list = [
    ['大きい', '小さい'],
    ['軽い', '重い'],
    ['親しみやすい', '親しみにくい'],
    ['大人っぽい', '子供っぽい'],
    ['柔らかい', '硬い'],
    ['フィットする', 'フィットしない'],
    ['通気性が良い', '通気性が悪い'],
    ['外れにくい', '外れやすい'],
    ['しゃべりやすい', 'しゃべりにくい'],
    ['暖かい', '寒い'],
    ['息がしやすい', '息がしにくい'],
    ['カジュアル', 'エレガント']
]

# Wordファイルを読み込む
doc = docx.Document(WORD_PASS)

for tbl in doc.tables:
    # 形容詞対リストをシャッフルする
    for row in adjective_pair_list:
        shuffle(row)
    shuffle(adjective_pair_list)

    for i in range(len(adjective_pair_list)):
        # 形容詞を変更する
        tbl.rows[i + 1].cells[0].text = adjective_pair_list[i][0]
        tbl.rows[i + 1].cells[2].text = adjective_pair_list[i][1]
        # 中央揃えにする
        tbl.rows[i + 1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        tbl.rows[i + 1].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # 列の幅を自動調節する
    tbl.autofit = True

# Wordファイルを保存する
doc.save(NEW_WORD_PASS)
