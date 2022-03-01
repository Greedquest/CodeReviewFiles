# Main download file
[TypeInfo.xlam](./TypeInfo.xlam?raw=True) is the code from the [CR post](https://codereview.stackexchange.com/questions/274532/low-level-vba-hacking-making-private-functions-public) packaged as an addin which you can add a reference to.
It requires adding references to the files below:

### These are the References
 
 - [MemoryTools.xlam](./MemoryTools.xlam?raw=True) - This is an addin which wraps [cristianbuse](https://github.com/cristianbuse)/**[VBA-MemoryTools](https://github.com/cristianbuse/VBA-MemoryTools)** which I'm using to read/write memory e.g. `MemByte(address As LongPtr) = value` because it is both performant and has a really nice API design in my opinion.
 - `TLBINF32.dll` -    This is a nice wrapper library for dealing with `ITypeLib` and `ITypeInfo` reflection* interfaces. However, it has some drawbacks:
	 - On 64-bit VBA it needs to be wrapped in a "COM+ server" since it is only a 32-bit library ([install instructions][3]).
	 - It is no longer shipped with Windows so has to be obtained from dodgy sites ([download][4]).
	 - More importantly, it cannot process the full `ITypeInfo` and filters out only the public members. As you will see this restricted the usage of this dll.
 - [COMTools.xlam](./COMTools.xlam?raw=True)  - This is an addin I wrote myself for this project and contains all the types and library functions to make working with COM possible in VBA. In particular:
	 - VTables** for `IUnknown`, `IDispatch` and the other various interfaces that crop up
	 - Standard methods like `ObjectFromObjPtr` and `QueryInterface` for dealing with interfaces
	 - Methods `CallFunction`, `CallCOMObjectVTableEntry` & `CallVBAFuncPtr` which wrap `DispCallFunc` and allow you to invoke function pointers
	 - _NOTE you must open this and add a reference to `MemoryTools.xlam` since it relies on that addin too_

**ALL vba prjects are password protected; password = "1"**
