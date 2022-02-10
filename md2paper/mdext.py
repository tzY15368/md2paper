from functools import reduce
from markdown.extensions import Extension
from markdown.inlinepatterns import SimpleTagPattern
from markdown.blockprocessors import BlockProcessor
import xml.etree.ElementTree as etree
import markdown
import re


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
                e.text = reduce(lambda x, y: x+'\n'+y, blocks[0:block_num + 1])
                #self.parser.parseBlocks(e, blocks[0:block_num + 1])
                # remove used blocks
                for i in range(0, block_num + 1):
                    blocks.pop(0)
                return True  # or could have had no return statement
        # No closing marker!  Restore and do nothing
        blocks[0] = original_block
        return False  # equivalent to our test() routine returning False


class MDExt(Extension):
    MATH_INLINE_RE = r'(\$)(.*?)\$'
    REF_RE = r'(\[)(.*?)\]'

    def extendMarkdown(self, md):
        # Create the del pattern
        math_inline_tag = SimpleTagPattern(self.MATH_INLINE_RE, 'math-inline')
        ref_tag = SimpleTagPattern(self.REF_RE, 'ref')
        # Insert del pattern into markdown parser
        md.inlinePatterns.register(math_inline_tag, 'math-inline', 75)
        md.inlinePatterns.register(ref_tag, 'ref', 75)

        md.parser.blockprocessors.register(
            MathBlockProcessor(md.parser), 'math', 75)


if __name__ == "__main__":
    md = '''
# a

$$这里假装是个公式$$

$$ \sum $$

$$
\sum
$$

inline $a$ formula

[引用]

    '''
    a = markdown.markdown(md, extensions=[MDExt()])
    print(a)
