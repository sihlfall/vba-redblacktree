# A red–black tree implementation for VBA / VB6

This is a translation of the red–black tree example implementation that can be found on [Wikipedia](https://en.wikipedia.org/wiki/Red%E2%80%93black_tree) to VBA. We support insertion only.

The nodes are kept in an array. `RedBlackInsert` inserts a prepared node into the tree at a given location. `RedBlackFind` searches for a node in the tree or—if the node does not exist—for the appropriate location where the corresponding node would have to be inserted.

The implementation is meant to serve as a template, which can be customized by adjusting the `NodeTypeTemplate` type and replacing the call to `RedBlackComparatorTemplate` in `RedBlackFind` by an appropriate comparison.

For an example how to use the code, have a look at subroutine `RunTest` in [`test/RedBlackTreeTest.bas`](test/RedBlackTreeTest.bas).