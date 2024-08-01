### Copyright (c) 2024 Bitech Systems. All rights reserved.
### The code and materials in this repository are the exclusive property of Bitech Systems and its associated companies and are protected by copyright law.
### Please refer to the license details in the package.


class TagIterator:
    def __init__(self, open_tag="[T_", close_tag="_T]"):
        self.open_tag = open_tag
        self.close_tag = close_tag
        self.initialized = False
        self.fnProcess = None

    def make_obj(self, level=0, text="", meta=None):
        self.lastid += 1
        return {
            "text": text,
            "level": level,
            "children": [],
            "meta": meta,
            "id": self.lastid,
        }

    #
    def start(self):
        # Initialize or reset the state
        self.lastid = 0
        self.initialized = True
        self.errors = []
        self.root = self.make_obj()
        self.current = self.root
        self.level = 0
        self.opens = 0
        self.closes = 0

    def process_text(self, text, meta=None):
        if not self.initialized:
            raise RuntimeError("You must call start() before process_text().")
        #

        def getNode(node, level):
            if node is None:
                return None
            if node.get("level", -1) == self.level:
                return node
            for child in node.get("children", []):
                result = getNode(child, level)
                if result is not None:
                    return result
                #
            #

        #

        i = 0
        lastParent = self.current
        while i < len(text):
            if text[i : i + len(self.open_tag)] == self.open_tag:
                self.level += 1
                self.opens += 1
                objref = self.make_obj(self.level, "", meta)
                self.current["text"] += "[{0}]".format(objref["id"])
                self.current["children"].append(objref)
                lastParent = self.current
                self.current = objref
                i += len(self.open_tag)

            elif text[i : i + len(self.close_tag)] == self.close_tag:
                if self.level > 0:
                    self.level -= 1
                #
                self.closes += 1
                self.current = getNode(self.root, self.level)
                i += len(self.close_tag)
            else:
                self.current["text"] += text[i]
                i += 1
            #
        #

    #

    def finish(self):
        if not self.initialized:
            raise RuntimeError("You must call start() before finish().")

        if self.closes != self.opens:
            raise AssertionError(
                "Closing ({}) and Opening ({}) tags does not match.".format(
                    self.closes, self.opens
                )
            )
        #
        return self.root

    def as_merge_tags(self, startID=1, parentID=0, isRoot=False):

        self.merge_tags_startid = startID

        def loopChildren(item, itemID, parentID):
            nodes = {}

            # if item.get("level",0) > 0 or isRoot:
            nodes["{},{}".format(itemID, parentID)] = {
                "text": item["text"],
                "meta": item["meta"],
            }
            #
            cnt = itemID
            if item.get("level", 0) == 0 and not isRoot:
                itemID = parentID
            #
            for child in item.get("children", []):
                cnt += 1
                childNodes = loopChildren(child, cnt, itemID)
                nodes.update(childNodes)
            #
            if cnt > self.merge_tags_startid:
                self.merge_tags_startid = startID
            #
            return nodes

        #

        flatlist = loopChildren(self.finish(), startID, parentID)
        if callable(self.fnProcess):
            tags = {}
            for key in flatlist:
                item = flatlist[key]
                tags[key] = self.fnProcess(item["text"])
            #
            return tags
        #
        return flatlist

    #
