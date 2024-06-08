#! /usr/bin/env python
#  -*- coding: utf-8 -*-

from __future__ import print_function, absolute_import
import sys, os
original_import = __import__

_indentation = 0


def _normalizePath(path):
    path = os.path.abspath(path)

    best = None

    for path_entry in sys.path:
        if path.startswith(path_entry):
            if best is None or len(path_entry) > len(best):
                best = path_entry

    if best is not None:
        path = path.replace(best, "$PYTHONPATH")

    return path


def _moduleRepr(module):
    try:
        module_file = module.__file__
        module_file = module_file.replace(".pyc", ".py")

        if module_file.endswith(".so"):
            module_file = os.path.join(
                os.path.dirname(module_file),
                os.path.basename(module_file).split(".")[0] + ".so",
            )

        file_desc = _normalizePath(module_file).replace(".pyc", ".py")
    except AttributeError as exc:
        file_desc = "built-in"
    return (module.__name__, file_desc)


def enableImportTracing(normalize_paths=True, show_source=False):
    def _ourimport(
        name,
        globals=None,
        locals=None,
        fromlist=None,  # @ReservedAssignment
        level=-1 if sys.version_info[0] < 3 else 0,
    ):
        builtins.__import__ = original_import

        global logfile
        global _indentation
        try:
            _indentation += 1

            logfile.write("%i;CALL;%s;%s\n" % (_indentation, name, fromlist))

            for entry in traceback.extract_stack()[:-1]:
                if entry[2] == "_ourimport":
                    continue
                else:
                    entry = list(entry)

                    if not show_source:
                        del entry[-1]
                        del entry[-1]

                    if normalize_paths:
                        entry[0] = _normalizePath(entry[0])

            builtins.__import__ = _ourimport
            try:
                result = original_import(name, globals, locals, fromlist, level)
            except ImportError as e:
                logfile.write("%i;EXCEPTION;%s\n" % (_indentation, e))
                result = None
                raise

            if result is not None:
                m = _moduleRepr(result)
                logfile.write("%i;RESULT;%s;%s\n" % (_indentation, m[0], m[1]))

            builtins.__import__ = _ourimport

            return result
        finally:
            _indentation -= 1

    try:
        import __builtin__ as builtins
    except ImportError:
        import builtins

    import traceback

    builtins.__import__ = _ourimport

scriptname = r"/Users/jd/dev/consulting/python_script_parser/scripti"
extname = ".py"
hinter_pid = "44454"
lname = "%s-%s-%s.log" % (scriptname, hinter_pid, os.getpid())  # each process has its logfile
logfile = open(lname, "w", buffering=1)
hints_logfile = logfile
source_file = open(scriptname + extname, encoding='utf-8')
source = source_file.read()
source_file.close()
enableImportTracing()
exec(source)
