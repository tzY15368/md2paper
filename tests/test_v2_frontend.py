import sys
sys.path.append('..')
from md2paper.v2.frontend import Paper, DUTPaperPreprocessor
import logging

logging.getLogger().setLevel(logging.DEBUG)

if __name__ == '__main__':
    p = Paper("../example/论文.md", DUTPaperPreprocessor)
