# A red–black tree implementation for VBA / VB6

This is a VBA implementation of a red–black tree. We support insertion only.

The nodes are kept in an array. `RedBlackInsert` inserts a prepared node into the tree at a given location. `RedBlackFindPosition` searches for a node in the tree; if the node does not exist in the tree, it provides the location at which the node would have to be inserted.

The implementation is meant to serve as a template, which can be customized by adjusting the `NodeTypeTemplate` type and replacing the call to `RedBlackComparatorTemplate` in `RedBlackFindPosition` by an appropriate comparison. For an example on how to use the code, have a look at the subroutine `RunTest` in [`test/RedBlackTreeTest.bas`](test/RedBlackTreeTest.bas).

An explanation of the algorithm can be found on [Wikipedia](https://en.wikipedia.org/w/index.php?title=Red%E2%80%93black_tree&oldid=1150140777) and [Wikibooks](https://en.wikibooks.org/w/index.php?title=F_Sharp_Programming/Advanced_Data_Structures&oldid=4052491). The variable names we use are similar to the ones used in the code examples there.