#!/usr/bin/env python
"""A LaTeX subset -> OMML, so slide text can carry real PowerPoint equations.

PowerPoint stores an inline equation as an ``mc:AlternateContent`` block inside
the paragraph: the ``a14:m`` choice holds the OMML that the equation editor
edits, and the fallback holds plain text for readers that do not understand it.
Writing ``L_rw`` as literal characters is not an equation -- it neither renders
with the right glyphs nor opens in the editor.

Supported: sub/superscripts, \\frac, \\sqrt, \\hat/\\bar/\\tilde/\\vec,
\\mathcal/\\mathbb/\\mathrm/\\text, greek letters, and the usual operators.
That covers what a paper's inline math needs; a full display-math typesetter is
out of scope.
"""
from __future__ import annotations

import re
from xml.sax.saxutils import escape

from pptx.oxml import parse_xml

NS = (
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
)

GREEK = {
    'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ', 'epsilon': 'ε',
    'varepsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ', 'iota': 'ι',
    'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ', 'nu': 'ν', 'xi': 'ξ', 'pi': 'π',
    'rho': 'ρ', 'sigma': 'σ', 'tau': 'τ', 'upsilon': 'υ', 'phi': 'φ',
    'varphi': 'φ', 'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',
    'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ', 'Xi': 'Ξ',
    'Pi': 'Π', 'Sigma': 'Σ', 'Phi': 'Φ', 'Psi': 'Ψ', 'Omega': 'Ω',
}

SYMBOL = {
    'times': '×', 'cdot': '⋅', 'div': '÷', 'pm': '±', 'mp': '∓',
    'leq': '≤', 'le': '≤', 'geq': '≥', 'ge': '≥', 'neq': '≠', 'ne': '≠',
    'approx': '≈', 'equiv': '≡', 'propto': '∝', 'sim': '∼',
    'rightarrow': '→', 'to': '→', 'leftarrow': '←', 'Rightarrow': '⇒',
    'leftrightarrow': '↔', 'mapsto': '↦',
    'sum': '∑', 'prod': '∏', 'int': '∫', 'partial': '∂', 'nabla': '∇',
    'infty': '∞', 'forall': '∀', 'exists': '∃', 'in': '∈', 'notin': '∉',
    'subset': '⊂', 'subseteq': '⊆', 'cup': '∪', 'cap': '∩', 'odot': '⊙',
    'oplus': '⊕', 'otimes': '⊗', 'circ': '∘', 'ldots': '…', 'cdots': '⋯',
    'langle': '⟨', 'rangle': '⟩', '|': '|', 'quad': ' ', 'qquad': '  ',
    ',': ' ', ';': ' ', ' ': ' ',
}

ACCENT = {'hat': '̂', 'bar': '̄', 'tilde': '̃', 'vec': '⃗',
          'dot': '̇', 'widehat': '̂', 'widetilde': '̃'}

# Unicode has one code point per styled letter; PowerPoint renders these
# directly, which is simpler and more portable than an m:rPr script style.
_CAL = {'B': 'ℬ', 'E': 'ℰ', 'F': 'ℱ', 'H': 'ℋ', 'I': 'ℐ', 'L': 'ℒ',
        'M': 'ℳ', 'R': 'ℛ', 'e': 'ℯ', 'g': 'ℊ', 'o': 'ℴ'}
_BB = {'C': 'ℂ', 'H': 'ℍ', 'N': 'ℕ', 'P': 'ℙ', 'Q': 'ℚ', 'R': 'ℝ', 'Z': 'ℤ'}


def _styled(text: str, table: dict, base: int) -> str:
    out = []
    for ch in text:
        if ch in table:
            out.append(table[ch])
        elif 'A' <= ch <= 'Z':
            out.append(chr(base + ord(ch) - ord('A')))
        elif 'a' <= ch <= 'z':
            out.append(chr(base + 26 + ord(ch) - ord('a')))
        else:
            out.append(ch)
    return ''.join(out)


class LatexError(ValueError):
    pass


def _tokenize(src: str) -> list[tuple[str, str]]:
    toks, i = [], 0
    while i < len(src):
        c = src[i]
        if c == '\\':
            m = re.match(r'\\([a-zA-Z]+|.)', src[i:])
            if not m:
                raise LatexError(f'dangling backslash in {src!r}')
            toks.append(('cmd', m.group(1)))
            i += m.end()
        elif c in '{}_^':
            toks.append((c, c))
            i += 1
        elif c.isspace():
            i += 1
        else:
            toks.append(('chr', c))
            i += 1
    return toks


def _run(text: str, size: float, upright: bool = False) -> str:
    """One math run. OMML italicises latin letters unless the style says plain."""
    sty = '<m:rPr><m:sty m:val="p"/></m:rPr>' if upright else ''
    rpr = (f'<a:rPr lang="en-US" sz="{int(round(size * 100))}">'
           f'<a:latin typeface="Cambria Math"/></a:rPr>')
    return f'{"<m:r>"}{sty}{rpr}<m:t xml:space="preserve">{escape(text)}</m:t></m:r>'


def _wrap(frags: list[str]) -> str:
    return ''.join(frags) if frags else ''


def _parse(toks: list[tuple[str, str]], i: int, size: float,
           until_brace: bool = False) -> tuple[list[str], int]:
    out: list[str] = []
    while i < len(toks):
        kind, val = toks[i]
        if kind == '}':
            if until_brace:
                return out, i + 1
            raise LatexError('unbalanced }')
        if kind in ('_', '^'):
            base = out.pop() if out else _run('', size)
            sub = sup = None
            while i < len(toks) and toks[i][0] in ('_', '^'):
                which = toks[i][0]
                arg, i = _argument(toks, i + 1, size)
                if which == '_':
                    sub = arg
                else:
                    sup = arg
            if sub is not None and sup is not None:
                out.append(f'<m:sSubSup><m:e>{base}</m:e><m:sub>{sub}</m:sub>'
                           f'<m:sup>{sup}</m:sup></m:sSubSup>')
            elif sub is not None:
                out.append(f'<m:sSub><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>')
            else:
                out.append(f'<m:sSup><m:e>{base}</m:e><m:sup>{sup}</m:sup></m:sSup>')
            continue
        if kind == '{':
            frags, i = _parse(toks, i + 1, size, until_brace=True)
            out.append(_wrap(frags))
            continue
        if kind == 'chr':
            out.append(_run(val, size))
            i += 1
            continue
        # kind == 'cmd'
        i += 1
        if val in ('frac', 'dfrac', 'tfrac'):
            num, i = _argument(toks, i, size)
            den, i = _argument(toks, i, size)
            out.append(f'<m:f><m:num>{num}</m:num><m:den>{den}</m:den></m:f>')
        elif val == 'sqrt':
            rad, i = _argument(toks, i, size)
            out.append('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>'
                       f'<m:deg/><m:e>{rad}</m:e></m:rad>')
        elif val in ACCENT:
            base, i = _argument(toks, i, size)
            out.append(f'<m:acc><m:accPr><m:chr m:val="{ACCENT[val]}"/></m:accPr>'
                       f'<m:e>{base}</m:e></m:acc>')
        elif val in ('mathcal', 'mathbb', 'mathrm', 'text', 'operatorname', 'mathbf'):
            raw, i = _raw_argument(toks, i)
            if val == 'mathcal':
                out.append(_run(_styled(raw, _CAL, 0x1D49C), size))
            elif val == 'mathbb':
                out.append(_run(_styled(raw, _BB, 0x1D538), size))
            elif val == 'mathbf':
                out.append(_run(_styled(raw, {}, 0x1D400), size))
            else:
                out.append(_run(raw, size, upright=True))
        elif val in ('left', 'right'):
            if i < len(toks):                       # the delimiter it applies to
                out.append(_run(toks[i][1].replace('.', ''), size))
                i += 1
        elif val in GREEK:
            out.append(_run(GREEK[val], size))
        elif val in SYMBOL:
            out.append(_run(SYMBOL[val], size, upright=True))
        else:
            raise LatexError(f'unsupported command \\{val}')
    if until_brace:
        raise LatexError('unbalanced {')
    return out, i


def _argument(toks, i, size) -> tuple[str, int]:
    """One argument: a braced group, or the single token that follows."""
    if i >= len(toks):
        raise LatexError('missing argument')
    if toks[i][0] == '{':
        frags, i = _parse(toks, i + 1, size, until_brace=True)
        return _wrap(frags), i
    frags, j = _parse(toks[i:i + 1], 0, size)
    return _wrap(frags), i + 1


def _raw_argument(toks, i) -> tuple[str, int]:
    """The literal characters of an argument -- for \\mathrm and friends."""
    if i >= len(toks) or toks[i][0] != '{':
        if i < len(toks) and toks[i][0] == 'chr':
            return toks[i][1], i + 1
        raise LatexError('missing argument')
    i += 1
    buf = []
    while i < len(toks) and toks[i][0] != '}':
        if toks[i][0] not in ('chr', 'cmd'):
            raise LatexError('nested markup not supported here')
        buf.append(toks[i][1])
        i += 1
    return ''.join(buf), i + 1


def to_plain(latex: str) -> str:
    """A readable one-line rendering, used for the fallback and for sizing."""
    s = latex
    for name, ch in {**GREEK, **SYMBOL}.items():
        s = s.replace('\\' + name, ch)
    s = re.sub(r'\\(?:frac|dfrac|tfrac)\{(.*?)\}\{(.*?)\}', r'(\1)/(\2)', s)
    s = re.sub(r'\\(?:mathcal|mathbb|mathrm|mathbf|text|operatorname|hat|bar|'
               r'tilde|vec|dot|sqrt|left|right)', '', s)
    return s.replace('{', '').replace('}', '').replace('\\', '')


def math_element(latex: str, size: float = 18):
    """An mc:AlternateContent holding `latex` as a PowerPoint equation."""
    frags, _ = _parse(_tokenize(latex), 0, size)
    omml = _wrap(frags)
    fallback = (f'<a:r><a:rPr lang="en-US" sz="{int(round(size * 100))}"/>'
                f'<a:t>{escape(to_plain(latex))}</a:t></a:r>')
    return parse_xml(
        f'<mc:AlternateContent {NS}>'
        f'<mc:Choice Requires="a14"><a14:m><m:oMath>{omml}</m:oMath></a14:m></mc:Choice>'
        f'<mc:Fallback>{fallback}</mc:Fallback>'
        f'</mc:AlternateContent>')
