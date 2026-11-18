(() => {
  "use strict";

  const typeMarker = /[\[［][ \u00a0]*(J\/OL|EB\/OL|[JDMCPNRSAGBKTOZ])[ \u00a0]*[\]］]/giu;
  const authorList = /^(?<authors>[\p{Script=Han}·]{2,8}(?:[ \u00a0]*[,，、][ \u00a0]*(?:[\p{Script=Han}·]{2,8}|等))*)[ \u00a0]*[.。．][ \u00a0]*(?=\S)/du;
  const nameWord = /^\p{Lu}[\p{Script=Latin}\p{M}'’‐‑-]+$/u;
  const initials = /^\p{Lu}\.$/u;
  const nameParticles = new Set(["de", "del", "da", "dos", "van", "von", "der", "ten"]);
  const journal = /^[ \u00a0]*(?:[.。．][ \u00a0]*)?(?<publication>\S(?:.*?\S)?)[ \u00a0]*[,，][ \u00a0]*(?<year>[12]\d{3})[ \u00a0]*[,，][ \u00a0]*(?<volume>\d{1,6})?[ \u00a0]*(?:[(（][ \u00a0]*(?<issue>[A-Za-z]?\d+[A-Za-z]?(?:[-–]\d+)?)[ \u00a0]*[)）])?[ \u00a0]*(?:[:：][ \u00a0]*(?<locator>[A-Za-z]?\d+[A-Za-z]?(?:[ \u00a0]*[-–—－][ \u00a0]*[A-Za-z]?\d+[A-Za-z]?)?))?[ \u00a0]*[.。．]?[ \u00a0]*$/du;
  const thesis = /^[ \u00a0]*(?:[.。．][ \u00a0]*)?(?:(?<advisorLabel>导师)[ \u00a0]*[:：][ \u00a0]*(?<advisor>[^.。．]+?)[ \u00a0]*[.。．][ \u00a0]*)?(?<institution>\S(?:.*?\S)?)[ \u00a0]*[,，][ \u00a0]*(?<year>[12]\d{3})[ \u00a0]*[.。．]?[ \u00a0]*$/du;

  function hasLatinAuthors(prefix) {
    for (const divider of prefix.matchAll(/[.。．][ \u00a0]+(?=\S)/gu)) {
      const names = prefix.slice(0, divider.index).split(/[,，]/u);
      if (names.every((author) => {
        const words = author.trim().split(/[ \u00a0]+/u);
        return words.length >= 2 && words.length <= 7 && nameWord.test(words.at(-1))
          && words.every((word) => nameWord.test(word) || initials.test(word) || nameParticles.has(word));
      })) return true;
    }
    return false;
  }

  function formatCitation(text) {
    if (text.length > 2500 || /[\r\n\t]/u.test(text)) {
      return { reason: "条目过长或含手动换行、制表符" };
    }

    const markers = [...text.matchAll(typeMarker)];
    if (markers.length !== 1) return { reason: "无法唯一确认文献类型" };
    const marker = markers[0];
    const type = marker[1].toUpperCase();
    if (type !== "J" && type !== "D") return { reason: "首版支持期刊 [J] 和学位论文 [D]" };

    const prefix = text.slice(0, marker.index).trimEnd();
    const authorMatch = authorList.exec(prefix);
    if (!authorMatch && !hasLatinAuthors(prefix)) {
      return { reason: "无法确认作者与题名的分隔" };
    }

    const markerEnd = marker.index + marker[0].length;
    const tail = text.slice(markerEnd);
    const match = (type === "J" ? journal : thesis).exec(tail);
    if (type === "J" && (!match || !(match.groups.volume || match.groups.issue))) {
      return { reason: "无法识别刊名、年份或卷期" };
    }
    if (type === "J" && !match.groups.locator) return { reason: "缺少页码或文章编号" };
    if (type === "D" && (!match || (tail.replace(/^[.。．\s]*/u, "").startsWith("导师") && !match.groups.advisor))) {
      return { reason: "无法识别学位授予单位、年份或导师字段" };
    }

    const edits = [];
    const edit = (start, end, replacement) => {
      if (text.slice(start, end) !== replacement) edits.push({ start, end, replacement });
    };
    if (authorMatch) {
      const [start, end] = authorMatch.indices.groups.authors;
      for (const divider of authorMatch.groups.authors.matchAll(/[ \u00a0]*[,，、][ \u00a0]*/gu)) {
        edit(start + divider.index, start + divider.index + divider[0].length, ", ");
      }
      edit(end, authorMatch[0].length, ". ");
    }
    edit(prefix.length, marker.index, "");
    edit(marker.index, markerEnd, "[" + type + "]");

    const group = (name) => match.indices.groups[name].map((index) => index + markerEnd);
    const between = (previous, next, separator) => edit(group(previous)[1], group(next)[0], separator);

    if (type === "J") {
      edit(markerEnd, group("publication")[0], ". ");
      between("publication", "year", ", ");
      if (match.groups.volume) between("year", "volume", ", ");
      if (match.groups.issue) between(match.groups.volume ? "volume" : "year", "issue", match.groups.volume ? "(" : ", (");
      between(match.groups.issue ? "issue" : "volume", "locator", match.groups.issue ? "): " : ": ");
      const range = /[ \u00a0]*[-–—－][ \u00a0]*/u.exec(match.groups.locator);
      if (range) {
        const start = group("locator")[0] + range.index;
        edit(start, start + range[0].length, "-");
      }
      edit(group("locator")[1], text.length, ".");
    } else {
      edit(markerEnd, group(match.groups.advisor ? "advisorLabel" : "institution")[0], ". ");
      if (match.groups.advisor) {
        between("advisorLabel", "advisor", ": ");
        between("advisor", "institution", ". ");
      }
      between("institution", "year", ", ");
      edit(group("year")[1], text.length, ".");
    }

    return { edits };
  }

  window.formatCitation = formatCitation;
})();
