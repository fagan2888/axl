import re, importlib, types, sys, os, glob

_imports = types.ModuleType('axl_imports')
_modules = {}


def add_symbol(module_name, symbol_name, value):
	idict = _imports.__dict__
	if symbol_name not in idict:
		idict[symbol_name] = value
		_modules[symbol_name] = module_name
	elif value is not idict[symbol_name]:
		other_module = _modules[symbol_name]
		other_symbol = '{}.{}'.format(other_module, symbol_name) if other_module else symbol_name
		symbol_name = '{}.{}'.format(module_name, symbol_name) if module_name else symbol_name
		raise ImportError('Conflict between symbol names:\n  - {}\n  - {}'.format(other_symbol, symbol_name))


def parse_input_line(line):
	line = line.split('#', 1)[0].strip()
	if not line:
		return
	module = symbols = None
	is_import = re.match('^import\s+(.*)', line)
	if is_import:
		module, symbols = is_import.groups() + (None,)
	else:
		is_from = re.match('^from\s+(\S+)\s+import\s+(.*)$', line)
		if is_from:
			module, symbols = is_from.groups()
		else:
			raise RuntimeError('Invalid import string: {}'.format(line))
	if module.startswith('.'):
		raise RuntimeError('Relative imports are not allowed: {}'.format(line))
	mod = importlib.import_module(module)
	if symbols is None:
		add_symbol(module.split('.', 1)[0], mod)
		return
	elif symbols != '*':
		symbols = re.split(',\s*', symbols)
	elif hasattr(mod, '__all__'):
		symbols = mod.__all__
	else:
		symbols = [name for name in dir(mod) if not name.startswith('_')]
	for name in symbols:
		if not hasattr(mod, name):
			raise ImportError('Cannot import name {!r} from module {!r}'.format(name, module))
		add_symbol(module, name, getattr(mod, name))


parse_input_line('from axl.methods import *')
import_glob = os.path.join(sys.prefix, 'Tools', 'axl', 'imports.*')
for fname in glob.glob(import_glob):
	with open(fname) as fp:
		for imp in fp.read().splitlines():
			parse_input_line(imp)




