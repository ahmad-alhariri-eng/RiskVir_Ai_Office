/**
 * Lightweight LaTeX → OMML converter for Word equations.
 * 
 * Supports: fractions, sqrt, superscript, subscript, Greek letters,
 *           operators, sum/int/prod with limits, text blocks, and nesting.
 *
 * The LLM generates LaTeX (which it excels at), and this converter
 * turns it into Office Math Markup Language (OMML) wrapped in OOXML,
 * ready for Word's insertOoxml() API.
 */

// ── Symbol maps ──────────────────────────────────────────────────

const GREEK: Record<string, string> = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε',
  varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ',
  iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν',
  xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ',
  upsilon: 'υ', phi: 'φ', varphi: 'ϕ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Alpha: 'Α', Beta: 'Β', Gamma: 'Γ', Delta: 'Δ', Epsilon: 'Ε',
  Zeta: 'Ζ', Eta: 'Η', Theta: 'Θ', Iota: 'Ι', Kappa: 'Κ',
  Lambda: 'Λ', Mu: 'Μ', Nu: 'Ν', Xi: 'Ξ', Pi: 'Π',
  Rho: 'Ρ', Sigma: 'Σ', Tau: 'Τ', Upsilon: 'Υ', Phi: 'Φ',
  Chi: 'Χ', Psi: 'Ψ', Omega: 'Ω',
  infty: '∞', partial: '∂', nabla: '∇', hbar: 'ℏ',
};

const SYMBOLS: Record<string, string> = {
  pm: '±', mp: '∓', times: '×', div: '÷', cdot: '⋅', ast: '∗',
  star: '⋆', circ: '∘', bullet: '∙',
  leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠',
  approx: '≈', equiv: '≡', sim: '∼', simeq: '≃', cong: '≅',
  ll: '≪', gg: '≫', propto: '∝',
  rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', Leftarrow: '⇐',
  leftrightarrow: '↔', Leftrightarrow: '⇔', to: '→', mapsto: '↦',
  uparrow: '↑', downarrow: '↓',
  'in': '∈', notin: '∉', ni: '∋',
  subset: '⊂', supset: '⊃', subseteq: '⊆', supseteq: '⊇',
  cup: '∪', cap: '∩', setminus: '∖', emptyset: '∅',
  forall: '∀', exists: '∃', nexists: '∄', neg: '¬',
  land: '∧', lor: '∨', oplus: '⊕', otimes: '⊗',
  ldots: '…', cdots: '⋯', vdots: '⋮', ddots: '⋱',
  perp: '⊥', parallel: '∥', angle: '∠', triangle: '△',
  ell: 'ℓ', Re: 'ℜ', Im: 'ℑ', wp: '℘', aleph: 'ℵ',
};

// ── XML helpers ──────────────────────────────────────────────────

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Wrap text in an OMML run element */
function run(text: string): string {
  return `<m:r><m:t>${esc(text)}</m:t></m:r>`;
}

// ── Parsing helpers ──────────────────────────────────────────────

/** Extract content of a brace group {…} starting at position i. Returns [content, newIndex]. */
function extractGroup(latex: string, i: number): [string, number] {
  while (i < latex.length && latex[i] === ' ') i++;
  if (i >= latex.length) return ['', i];

  if (latex[i] === '{') {
    let depth = 1;
    const start = i + 1;
    i++;
    while (i < latex.length && depth > 0) {
      if (latex[i] === '{') depth++;
      else if (latex[i] === '}') depth--;
      i++;
    }
    return [latex.slice(start, i - 1), i];
  }

  // Single character or single command as a "group"
  if (latex[i] === '\\') {
    let cmd = '\\';
    i++;
    while (i < latex.length && /[a-zA-Z]/.test(latex[i])) { cmd += latex[i]; i++; }
    return [cmd, i];
  }

  return [latex[i], i + 1];
}

// ── Main recursive parser ────────────────────────────────────────

function parse(latex: string): string {
  const elements: string[] = [];
  let i = 0;

  while (i < latex.length) {
    const ch = latex[i];

    if (ch === '\\') {
      // ── LaTeX command ──
      let cmd = '';
      i++;
      while (i < latex.length && /[a-zA-Z]/.test(latex[i])) { cmd += latex[i]; i++; }

      if (cmd === 'frac') {
        const [num, ni] = extractGroup(latex, i);
        const [den, di] = extractGroup(latex, ni);
        i = di;
        elements.push(`<m:f><m:num>${parse(num)}</m:num><m:den>${parse(den)}</m:den></m:f>`);

      } else if (cmd === 'sqrt') {
        // Check for optional n-th root: \sqrt[n]{x}
        let nthRoot = '';
        let j = i;
        while (j < latex.length && latex[j] === ' ') j++;
        if (j < latex.length && latex[j] === '[') {
          j++;
          const end = latex.indexOf(']', j);
          if (end !== -1) { nthRoot = latex.slice(j, end); j = end + 1; }
        }
        const [content, ni] = extractGroup(latex, j);
        i = ni;
        if (nthRoot) {
          elements.push(`<m:rad><m:deg>${parse(nthRoot)}</m:deg><m:e>${parse(content)}</m:e></m:rad>`);
        } else {
          elements.push(`<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${parse(content)}</m:e></m:rad>`);
        }

      } else if (cmd === 'sum' || cmd === 'prod' || cmd === 'int' ||
                 cmd === 'iint' || cmd === 'iiint' || cmd === 'oint') {
        const sym = cmd === 'sum' ? '∑' : cmd === 'prod' ? '∏' :
                    cmd === 'oint' ? '∮' : cmd === 'iint' ? '∬' :
                    cmd === 'iiint' ? '∭' : '∫';
        let sub = '', sup = '';
        let j = i;
        while (j < latex.length && latex[j] === ' ') j++;
        if (j < latex.length && latex[j] === '_') {
          j++;
          const [s, ni] = extractGroup(latex, j); sub = parse(s); j = ni;
        }
        while (j < latex.length && latex[j] === ' ') j++;
        if (j < latex.length && latex[j] === '^') {
          j++;
          const [s, ni] = extractGroup(latex, j); sup = parse(s); j = ni;
        }
        i = j;
        // Take next group as the body (if available)
        while (i < latex.length && latex[i] === ' ') i++;
        let body = '';
        if (i < latex.length && latex[i] === '{') {
          const [b, ni] = extractGroup(latex, i); body = parse(b); i = ni;
        } else if (i < latex.length) {
          // Take the rest of the expression as the body
          body = parse(latex.slice(i)); i = latex.length;
        }
        elements.push(`<m:nary><m:naryPr><m:chr m:val="${sym}"/></m:naryPr><m:sub>${sub || run('')}</m:sub><m:sup>${sup || run('')}</m:sup><m:e>${body}</m:e></m:nary>`);

      } else if (cmd === 'lim') {
        let sub = '';
        let j = i;
        while (j < latex.length && latex[j] === ' ') j++;
        if (j < latex.length && latex[j] === '_') {
          j++;
          const [s, ni] = extractGroup(latex, j); sub = parse(s); j = ni;
        }
        i = j;
        elements.push(`<m:func><m:fName>${run('lim')}</m:fName><m:e>${sub ? `<m:sSub><m:e>${run('')}</m:e><m:sub>${sub}</m:sub></m:sSub>` : run('')}</m:e></m:func>`);

      } else if (cmd === 'text' || cmd === 'mathrm' || cmd === 'textrm') {
        const [content, ni] = extractGroup(latex, i); i = ni;
        elements.push(`<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${esc(content)}</m:t></m:r>`);

      } else if (cmd === 'mathbf' || cmd === 'textbf' || cmd === 'boldsymbol') {
        const [content, ni] = extractGroup(latex, i); i = ni;
        elements.push(`<m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>${esc(content)}</m:t></m:r>`);

      } else if (cmd === 'overline' || cmd === 'bar') {
        const [content, ni] = extractGroup(latex, i); i = ni;
        elements.push(`<m:acc><m:accPr><m:chr m:val="̄"/></m:accPr><m:e>${parse(content)}</m:e></m:acc>`);

      } else if (cmd === 'hat') {
        const [content, ni] = extractGroup(latex, i); i = ni;
        elements.push(`<m:acc><m:accPr><m:chr m:val="̂"/></m:accPr><m:e>${parse(content)}</m:e></m:acc>`);

      } else if (cmd === 'vec') {
        const [content, ni] = extractGroup(latex, i); i = ni;
        elements.push(`<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr><m:e>${parse(content)}</m:e></m:acc>`);

      } else if (cmd === 'dot') {
        const [content, ni] = extractGroup(latex, i); i = ni;
        elements.push(`<m:acc><m:accPr><m:chr m:val="̇"/></m:accPr><m:e>${parse(content)}</m:e></m:acc>`);

      } else if (cmd === 'ddot') {
        const [content, ni] = extractGroup(latex, i); i = ni;
        elements.push(`<m:acc><m:accPr><m:chr m:val="̈"/></m:accPr><m:e>${parse(content)}</m:e></m:acc>`);

      } else if (cmd === 'tilde') {
        const [content, ni] = extractGroup(latex, i); i = ni;
        elements.push(`<m:acc><m:accPr><m:chr m:val="̃"/></m:accPr><m:e>${parse(content)}</m:e></m:acc>`);

      } else if (cmd === 'left') {
        // Collect until matching \right
        let j = i;
        if (j < latex.length) {
          const openDelim = latex[j] === '.' ? '' : latex[j]; j++;
          // Find matching \right
          let depth = 1;
          const bodyStart = j;
          while (j < latex.length && depth > 0) {
            if (latex.slice(j).startsWith('\\left')) { depth++; j += 5; }
            else if (latex.slice(j).startsWith('\\right')) { depth--; if (depth === 0) break; j += 6; }
            else j++;
          }
          const bodyContent = latex.slice(bodyStart, j);
          j += 6; // skip \right
          let closeDelim = '';
          if (j < latex.length) { closeDelim = latex[j] === '.' ? '' : latex[j]; j++; }
          i = j;
          const beginChr = openDelim ? ` <m:begChr m:val="${esc(openDelim)}"/>` : '<m:begChr m:val=""/>';
          const endChr = closeDelim ? ` <m:endChr m:val="${esc(closeDelim)}"/>` : '<m:endChr m:val=""/>';
          elements.push(`<m:d><m:dPr>${beginChr}${endChr}</m:dPr><m:e>${parse(bodyContent)}</m:e></m:d>`);
        }

      } else if (cmd === 'right') {
        // Should be handled by \left; skip
        if (i < latex.length) i++;

      } else if (cmd === 'binom') {
        const [top, ni] = extractGroup(latex, i);
        const [bot, di] = extractGroup(latex, ni);
        i = di;
        elements.push(`<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>${parse(top)}</m:num><m:den>${parse(bot)}</m:den></m:f></m:e></m:d>`);

      } else if (cmd === 'log' || cmd === 'ln' || cmd === 'sin' || cmd === 'cos' ||
                 cmd === 'tan' || cmd === 'sec' || cmd === 'csc' || cmd === 'cot' ||
                 cmd === 'arcsin' || cmd === 'arccos' || cmd === 'arctan' ||
                 cmd === 'sinh' || cmd === 'cosh' || cmd === 'tanh' ||
                 cmd === 'exp' || cmd === 'det' || cmd === 'max' || cmd === 'min' ||
                 cmd === 'sup' || cmd === 'inf' || cmd === 'gcd' || cmd === 'lcm' ||
                 cmd === 'deg' || cmd === 'dim' || cmd === 'ker' || cmd === 'mod') {
        // Math function — render as upright text
        elements.push(`<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${cmd}</m:t></m:r>`);

      } else if (GREEK[cmd]) {
        elements.push(run(GREEK[cmd]));

      } else if (SYMBOLS[cmd]) {
        elements.push(run(SYMBOLS[cmd]));

      } else if (cmd === '' && i < latex.length && !(/[a-zA-Z]/.test(latex[i]))) {
        // Escaped special character like \{ \} \% \#
        elements.push(run(latex[i]));
        i++;

      } else {
        // Unknown command — output as text
        elements.push(run(cmd));
      }

    } else if (ch === '^') {
      // ── Superscript ──
      i++;
      const [exp, ni] = extractGroup(latex, i); i = ni;
      // Check if followed by _ (sub-superscript)
      let j = i; while (j < latex.length && latex[j] === ' ') j++;
      if (j < latex.length && latex[j] === '_') {
        j++;
        const [sub, si] = extractGroup(latex, j); i = si;
        const base = elements.length > 0 ? elements.pop()! : run('');
        elements.push(`<m:sSubSup><m:e>${base}</m:e><m:sub>${parse(sub)}</m:sub><m:sup>${parse(exp)}</m:sup></m:sSubSup>`);
      } else {
        const base = elements.length > 0 ? elements.pop()! : run('');
        elements.push(`<m:sSup><m:e>${base}</m:e><m:sup>${parse(exp)}</m:sup></m:sSup>`);
      }

    } else if (ch === '_') {
      // ── Subscript ──
      i++;
      const [sub, ni] = extractGroup(latex, i); i = ni;
      // Check if followed by ^ (sub-superscript)
      let j = i; while (j < latex.length && latex[j] === ' ') j++;
      if (j < latex.length && latex[j] === '^') {
        j++;
        const [sup, si] = extractGroup(latex, j); i = si;
        const base = elements.length > 0 ? elements.pop()! : run('');
        elements.push(`<m:sSubSup><m:e>${base}</m:e><m:sub>${parse(sub)}</m:sub><m:sup>${parse(sup)}</m:sup></m:sSubSup>`);
      } else {
        const base = elements.length > 0 ? elements.pop()! : run('');
        elements.push(`<m:sSub><m:e>${base}</m:e><m:sub>${parse(sub)}</m:sub></m:sSub>`);
      }

    } else if (ch === '{') {
      // ── Brace group ──
      const [content, ni] = extractGroup(latex, i); i = ni;
      elements.push(parse(content));

    } else if (ch === ' ' || ch === '\n' || ch === '\t') {
      i++; // Skip whitespace in math mode

    } else {
      // ── Regular character (variable, number, operator) ──
      elements.push(run(ch));
      i++;
    }
  }

  return elements.join('');
}

// ── Public API ───────────────────────────────────────────────────

/**
 * Convert LaTeX math notation to a complete OOXML package
 * ready for Word's `insertOoxml()` API.
 */
export function latexToOoxml(latex: string): string {
  const omml = parse(latex.trim());
  return [
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">',
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">',
    '<pkg:xmlData>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">',
    '<w:body><w:p>',
    '<m:oMathPara><m:oMath>',
    omml,
    '</m:oMath></m:oMathPara>',
    '</w:p></w:body>',
    '</w:document>',
    '</pkg:xmlData>',
    '</pkg:part>',
    '</pkg:package>',
  ].join('');
}
