from functools import reduce
import markdown
from markdown.extensions import Extension
from markdown.inlinepatterns import InlineProcessor
from markdown.blockprocessors import BlockProcessor
from markdown.util import AtomicString
import xml.etree.ElementTree as etree
from xml.etree.ElementTree import Element
import re
from typing import *

from markdown.inlinepatterns import SimpleTagPattern


class MathBlockProcessor(BlockProcessor):
    RE_FENCE_START = r'^ *\$\$ *'  # start line, e.g., `$$`
    RE_FENCE_END = r' *\$\$$'  # last non-blank line, e.g, '$$'

    def test(self, parent, block):
        return re.match(self.RE_FENCE_START, block)

    def run(self, parent, blocks):
        original_block = blocks[0]
        blocks[0] = re.sub(self.RE_FENCE_START, '', blocks[0])
        # Find block with ending fence
        for block_num, block in enumerate(blocks):
            if re.search(self.RE_FENCE_END, block):
                # remove fence
                blocks[block_num] = re.sub(self.RE_FENCE_END, '', block)
                # render fenced area inside a new div
                e = etree.SubElement(parent, 'math')
                e.text = AtomicString(
                    reduce(lambda x, y: x+'\n'+y, blocks[0:block_num + 1]))
                #self.parser.parseBlocks(e, blocks[0:block_num + 1])
                # remove used blocks
                for i in range(0, block_num + 1):
                    blocks.pop(0)
                return True  # or could have had no return statement
        # No closing marker!  Restore and do nothing
        blocks[0] = original_block
        return False  # equivalent to our test() routine returning False


class RawTextPattern(InlineProcessor):
    def __init__(self, pattern, tag: str) -> None:
        super().__init__(pattern)
        self._tag = tag

    def handleMatch(self, m, data):
        node = Element(self._tag)
        # Text should not be further processed.
        node.text = AtomicString(m.group(2))
        return node, m.start(0), m.end(0)


class MDExt(Extension):
    def extendMarkdown(self, md):
        ref_tag = SimpleTagPattern(r'(\[)(.*?)\]', 'ref')
        md.inlinePatterns.register(ref_tag, 'ref', 75)

        md.inlinePatterns.register(
            RawTextPattern(r'(?<!\\|\$)(\$)([^\$]+)(\$)',  # $...$
                           'math-inline'),
            'math-inline',
            185)

        md.parser.blockprocessors.register(
            MathBlockProcessor(md.parser),
            'math',
            75)

        md.ESCAPED_CHARS.append('$')


if __name__ == "__main__":
    md = '''
# a

$$这里假装是个公式$$

$$ \sum $$

$$
\sum
$$

inline $a$ formula

[\{引用]

[a](b)

![a](b)

    '''
    a = markdown.markdown(md, extensions=[MDExt()])
    print(a)
