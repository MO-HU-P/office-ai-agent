"""段落内の一部の文字(run)だけに書式を付けるための共通処理。

Word(python-docx)もPowerPoint(python-pptx)も、段落は「run」(同じ書式が続く文字のまとまり)の
並びでできており、キーワードだけに色を付けるにはキーワードの境界でrunを分割して、
該当部分のrunにだけ書式を適用する必要がある。分割の手順は両者で同じなのでここに共通化する
(XML要素は w:r / a:r で異なるが、どちらも .text の取得/設定と lxml の addnext が使える)。

注意: Wordは編集履歴やスペルチェックの痕跡で、同じ書式でもテキストが複数のrunに
不規則に分かれていることがある。runごとに検索すると「ファインチューニング」が
「ファインチ」+「ューニング」に分かれていて見つからない、ということが起きるため、
段落全体のテキストで検索してから文字位置をrunに対応付ける方式をとる。
"""
from copy import deepcopy


def find_keyword_spans(text: str, keywords: list[str]) -> list[tuple[int, int]]:
    """テキスト内の全キーワードの出現範囲 (開始, 終了) を求める。重なる範囲は1つにまとめる。"""
    spans: list[tuple[int, int]] = []
    for kw in keywords:
        if not kw:
            continue
        start = 0
        while True:
            i = text.find(kw, start)
            if i < 0:
                break
            spans.append((i, i + len(kw)))
            start = i + len(kw)
    if not spans:
        return []
    spans.sort()
    merged = [spans[0]]
    for s, e in spans[1:]:
        if s <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    return merged


def split_runs_at_spans(run_els: list, spans: list[tuple[int, int]]) -> list:
    """runのXML要素の並びをspans(段落テキスト内の文字位置)の境界で分割し、
    span内に収まったrun要素のリストを返す。分割で増えたrunは元のrunの複製なので書式を引き継ぐ。
    呼び出し側は返された要素に対応するrunにだけ書式を適用すればよい。"""
    boundaries = sorted({b for span in spans for b in span})
    hit_els = []
    pos = 0
    for el in run_els:
        text = el.text or ""
        start, end = pos, pos + len(text)
        pos = end
        if not text:
            continue
        cuts = [b - start for b in boundaries if start < b < end]
        segs = []
        prev = 0
        for c in cuts + [len(text)]:
            segs.append(text[prev:c])
            prev = c
        seg_els = [el]
        for seg in segs[1:]:
            new_el = deepcopy(el)  # rPr(書式)ごと複製し、テキストだけ差し替える
            seg_els[-1].addnext(new_el)
            new_el.text = seg
            seg_els.append(new_el)
        if len(segs) > 1:
            el.text = segs[0]
        seg_pos = start
        for seg_el, seg in zip(seg_els, segs):
            if any(s <= seg_pos and seg_pos + len(seg) <= e for s, e in spans):
                hit_els.append(seg_el)
            seg_pos += len(seg)
    return hit_els


def validate_format_args(
    keywords, color, bold, italic, underline, font_size=None, font="", highlight=""
) -> tuple[list[str], str | None]:
    """3ツール共通の引数チェック。(正規化したキーワード, エラーメッセージ or None) を返す。"""
    kws = [k for k in (keywords or []) if isinstance(k, str) and k.strip()]
    if not kws:
        return [], "エラー: keywords(書式を付ける語句のリスト)を指定してください"
    if (not color and bold is None and italic is None and underline is None
            and font_size is None and not font and not highlight):
        return [], ("エラー: color / bold / italic / underline / font_size / font / highlight "
                    "のうち少なくとも1つを指定してください")
    for name, value, example in (("color", color, '赤字は "#FF0000"'), ("highlight", highlight, '黄色は "#FFFF00"')):
        if value:
            hexv = value.lstrip("#")
            if len(hexv) != 6 or any(c not in "0123456789abcdefABCDEF" for c in hexv):
                return [], f'エラー: {name}は "#RRGGBB" 形式で指定してください (例: {example})'
    if font_size is not None:
        try:
            ok = 1 <= float(font_size) <= 400
        except (TypeError, ValueError):
            ok = False
        if not ok:
            return [], "エラー: font_size(文字サイズ)は1〜400のポイント数で指定してください"
    return kws, None


def describe_format(color, bold, italic, underline, font_size=None, font="", highlight="") -> str:
    """適用した書式の報告用の説明文(例: "色#FF0000・太字")を作る。"""
    parts = []
    if color:
        parts.append(f"色{color if color.startswith('#') else '#' + color}")
    if bold is not None:
        parts.append("太字" if bold else "太字解除")
    if italic is not None:
        parts.append("斜体" if italic else "斜体解除")
    if underline is not None:
        parts.append("下線" if underline else "下線解除")
    if font_size is not None:
        parts.append(f"サイズ{float(font_size):g}pt")
    if font:
        parts.append(f"フォント{font}")
    if highlight:
        parts.append(f"蛍光ペン{highlight if highlight.startswith('#') else '#' + highlight}")
    return "・".join(parts)
