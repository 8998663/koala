# cython: profile=True

import os.path
from math import *
import json
import re
import networkx

from koala.openpyxl.formula.translate import Translator
from koala.reader import read_archive, read_named_ranges, read_cells
from koala.ExcelError import *
from koala.excellib import *
from koala.utils import *
from koala.ast import *
from koala.Cell import Cell
from koala.Range import RangeCore, RangeFactory, parse_cell_address, get_cell_address
from koala.tokenizer import reverse_rpn
from koala.serializer import *

class Spreadsheet(object):
    # You can create a spreadsheet in two ways:
    #   1) Through an excel file:
    #     sp = Spreadsheet()
    #     sp.loadExcel(filename, ignore_sheets = [], ignore_hidden = False)
    #   2) Through a gzip file containing a precalculated graph
    #     sp = Spreadsheet()
    #     sp.loadGraph()
    def __init__(self, cells, named_ranges, G = None, volatiles = set(), inputs = set(), outputs = set(), debug = False):
        self.cells = cells
        self.named_ranges = named_ranges
        self.G = G
        self.volatiles = volatiles
        self.inputs = outputs
        self.outputs = inputs
        self.debug = debug 

        self.volatile_to_remove = ["INDEX", "OFFSET"]
        self.Range = RangeFactory(cells)
        
        self.save_history = False
        self.history = dict()
        self.count = 0
        
        
        self.reset_buffer = set()
        self.fixed_cells = {}  

    @staticmethod
    def load_excel(file, ignore_sheets = [], ignore_hidden = False, debug = False):
        file_name = os.path.abspath(file)
        # Decompose subfiles structure in zip file
        archive = read_archive(file_name)
        # Parse cells
        cells = read_cells(archive, ignore_sheets, ignore_hidden)
        # Parse named_range { name (ExampleName) -> address (Sheet!A1:A10)}
        named_ranges = read_named_ranges(archive)

        return  Spreadsheet(cells, named_ranges, debug = debug)

    @staticmethod
    def load_gzip(fname, debug = False):
        return Spreadsheet(*load(fname), debug = debug)

    @staticmethod
    def load_json(fname, debug = False):
        return Spreadsheet(*load_json(fname), debug = debug)

    def dump_json(self, fname):
        dump_json(self, fname)

    def dump(self, fname):
        dump(self, fname)

    def activate_history(self):
        self.save_history = True

    def add_cell(self, cell, value = None):
        
        if type(cell) != Cell:
            cell = Cell(cell, None, value = value, formula = None, is_range = False, is_named_range = False)
        
        addr = cell.address()
        if addr in self.cells:
            raise Exception('Cell %s already in cellmap' % addr)

        cellmap, G = graph_from_seeds([cell], self, update = True)

        self.cells = cellmap
        self.G = G

        print "Graph construction updated, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))

    def set_formula(self, addr, formula):
        if addr in self.cells:
            cell = self.cells[addr]
        else:
            raise Exception('Cell %s not in cellmap' % addr)

        seeds = [cell]

        if cell.is_range:
            for index, c in enumerate(cell.range.cells): # for each cell of the range, translate the formula
                if index == 0:
                    c.formula = formula
                    translator = Translator(unicode('=' + formula), c.address().split('!')[1]) # the Translator needs a reference without sheet
                else:
                    translated = translator.translate_formula(c.address().split('!')[1]) # the Translator needs a reference without sheet
                    c.formula = translated[1:] # to get rid of the '='

                seeds.append(c)
        else:
            cell.formula = formula

        cellmap, G = graph_from_seeds(seeds, self)

        self.cells = cellmap
        self.G = G

        should_eval = self.cells[addr].should_eval
        self.cells[addr].should_eval = 'always'
        self.evaluate(addr)
        self.cells[addr].should_eval = should_eval

        print "Graph construction updated, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))


    def prune_graph(self):
        print '___### Pruning Graph ###___'

        G = self.G

        # get all the cells impacted by inputs
        dependencies = set()
        for input_address in self.inputs:
            child = self.cells[input_address]
            if child == None:
                print "Not found ", input_address
                continue
            g = make_subgraph(G, child, "descending")
            dependencies = dependencies.union(g.nodes())

        # print "%s cells depending on inputs" % str(len(dependencies))

        # prune the graph and set all cell independent of input to const
        subgraph = networkx.DiGraph()
        new_cellmap = {}
        for output_address in self.outputs:
            new_cellmap[output_address] = self.cells[output_address]
            seed = self.cells[output_address]
            todo = map(lambda n: (seed,n), G.predecessors(seed))
            done = set(todo)

            while len(todo) > 0:
                current, pred = todo.pop()
                # print "==========================="
                # print current.address(), pred.address()
                if current in dependencies:
                    if pred in dependencies or isinstance(pred.value, RangeCore) or pred.is_named_range:
                        subgraph.add_edge(pred, current)
                        new_cellmap[pred.address()] = pred
                        new_cellmap[current.address()] = current

                        nexts = G.predecessors(pred)
                        for n in nexts:            
                            if (pred,n) not in done:
                                todo += [(pred,n)]
                                done.add((pred,n))
                    else:
                        if pred.address() not in new_cellmap:
                            const_node = Cell(pred.address(), pred.sheet, value = pred.range if pred.is_range else pred.value, formula=None, is_range = isinstance(pred.range, RangeCore), is_named_range=pred.is_named_range, should_eval=pred.should_eval)
                            # pystr,ast = cell2code(self.named_ranges, const_node, pred.sheet)
                            # const_node.python_expression = pystr
                            # const_node.compile()
                            new_cellmap[pred.address()] = const_node

                        const_node = new_cellmap[pred.address()]
                        subgraph.add_edge(const_node, current)
                        
                else:
                    # case of range independant of input, we add all children as const
                    if pred.address() not in new_cellmap:
                        const_node = Cell(pred.address(), pred.sheet, value = pred.range if pred.is_range else pred.value, formula=None, is_range = pred.is_range, is_named_range=pred.is_named_range, should_eval=pred.should_eval)
                        # pystr,ast = cell2code(self.named_ranges, const_node, pred.sheet)
                        # const_node.python_expression = pystr
                        # const_node.compile()
                        new_cellmap[pred.address()] = const_node

                    const_node = new_cellmap[pred.address()]
                    subgraph.add_edge(const_node, current)


        print "Graph pruning done, %s nodes, %s edges, %s cellmap entries" % (len(subgraph.nodes()),len(subgraph.edges()),len(new_cellmap))

        # add back inputs that have been pruned because they are outside of calculation chain
        for i in self.inputs:
            if i not in new_cellmap:
                if i in self.named_ranges:
                    reference = self.named_ranges[i]
                    if is_range(reference):

                        rng = self.Range(reference)
                        virtual_cell = Cell(i, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                        new_cellmap[i] = virtual_cell
                        subgraph.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain

                    else:
                        # might need to be changed to actual self.cells Cell, not a copy
                        virtual_cell = Cell(i, None, value = self.cells[reference].value, formula = reference, is_range = False, is_named_range = True)
                        new_cellmap[i] = virtual_cell
                        subgraph.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain
                else:
                    if is_range(i):
                        rng = self.Range(i)
                        virtual_cell = Cell(i, None, value = rng, formula = o, is_range = True, is_named_range = True )
                        new_cellmap[i] = virtual_cell
                        subgraph.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain
                    else:
                        new_cellmap[i] = self.cells[i]
                        subgraph.add_node(self.cells[i]) # edges are not needed here since the input here is not in the calculation chain


        self.G = subgraph
        self.cells = new_cellmap

    def fix_volatiles(self, volatiles_to_fix):
        print '___### Fix Volatiles ###___'

        new_named_ranges = {}
        new_cells = {}

        # ### 1) create ranges
        # for n in self.named_ranges:
        #     reference = self.named_ranges[n]
        #     if is_range(reference):
        #         if 'OFFSET' not in reference:
        #             my_range = self.Range(reference)
        #             self.cells[n] = Cell(n, None, value = my_range, formula = reference, is_range = True, is_named_range = True )
        #         else:
        #             self.cells[n] = Cell(n, None, value = None, formula = reference, is_range = False, is_named_range = True )
        #     else:
        #         if reference in self.cells:
        #             self.cells[n] = Cell(n, None, value = self.cells[reference].value, formula = reference, is_range = False, is_named_range = True )
        #         else:
        #             self.cells[n] = Cell(n, None, value = None, formula = reference, is_range = False, is_named_range = True )
        
        # remove all volatile ranges from cells because they will refer to unnecessary nodes after
        # the ones we still need will be recreated during graph generation
        new_volatiles = self.volatiles.copy()
        for volatile_name in self.volatiles:
            volatile_cell = self.cells[volatile_name]
            if volatile_cell.is_range:
                new_volatiles.remove(volatile_name)
                del self.cells[volatile_name]
                try:
                    volatiles_to_fix.remove((volatile_cell.formula, volatile_cell.address(), volatile_cell.sheet))
                except:
                    pass
        for formula, address, sheet in volatiles_to_fix:


            new_volatiles.remove(address)

            if sheet:
                parsed = parse_cell_address(address)
            else:
                try:
                    successor = self.G.successors(self.cells[address])[0]
                    parsed = parse_cell_address(successor.address())
                    sheet = successor.sheet
                except Exception as e:
                    print e
                    parsed = ""
            e = shunting_yard(formula, self.named_ranges, ref = parsed, tokenize_range = True)
            ast,root = build_ast(e)
            # code = root.emit(ast)
            cell = {"formula": formula, "address": address, "sheet": sheet}
            replacements = self.eval_volatiles_from_ast(ast, root, cell)


            new_formula = formula
            if type(replacements) == list:
                for repl in replacements:
                    if type(repl["value"]) == ExcelError:
                        if self.debug:
                            print 'WARNING: Excel error found => replacing with #N/A'
                        repl["value"] = "#N/A"

                    if repl["expression_type"] == "value":
                        new_formula = new_formula.replace(repl["formula"], str(repl["value"]))
                    else:
                        new_formula = new_formula.replace(repl["formula"], repl["value"])
            else:
                new_formula = None

            if False:
                print 'Old', formula
                print 'new', new_formula


            if address in new_named_ranges:
                new_named_ranges[address] = new_formula
            else: 
                old_cell = self.cells[address]
                if old_cell.is_range:
                    cell = Cell(old_cell.address(), old_cell.sheet, value=old_cell.range, formula=new_formula, is_range = old_cell.is_range, is_named_range=old_cell.is_named_range, should_eval=old_cell.should_eval)
                else:
                    cell = Cell(old_cell.address(), old_cell.sheet, value=old_cell.value, formula=new_formula, is_range = old_cell.is_range, is_named_range=old_cell.is_named_range, should_eval=old_cell.should_eval)
                pystr, ast = cell2code(cell, self.named_ranges, tokenize_range = True)
                cell.python_expression = pystr.replace('"', "'")
                cell.compile()
                new_cells[address] = cell

        return new_cells, new_named_ranges, new_volatiles


    def print_value_ast(self, ast,node,indent):
        print "%s %s %s %s" % (" "*indent, str(node.token.tvalue), str(node.token.ttype), str(node.token.tsubtype))
        for c in node.children(ast):
            self.print_value_ast(ast, c, indent+1)

    def eval_volatiles_from_ast(self, ast, node, cell, debug = False):
        results = []
        context = cell["sheet"]

        if debug:
            print "tvalue ",node.token.tvalue

        if (node.token.tvalue == "INDEX" or node.token.tvalue == "OFFSET"):
            volatile_string = reverse_rpn(node, ast)
            expression = node.emit(ast, context=context)

            if expression.startswith("self.eval_ref"):
                expression_type = "value"
            else:
                expression_type = "formula"
            
            try:
                volatile_value = eval(expression)
            except Exception as e:
                if self.debug:
                    print RangeCore.apply("add",RangeCore.apply("substract",self.eval_ref("E70", ref = (122, 'I')),self.eval_ref("year_modelStart", ref = (122, 'I')),(122, 'I')),1,(122, 'I'))
                    print 'EXCEPTION raised in eval_volatiles: EXPR', expression, cell["address"]
                raise Exception("Problem evalling: %s for %s, %s" % (e, cell["address"], expression))

            return {"formula":volatile_string, "value": volatile_value, "expression_type": expression_type}      
        else:
            for c in node.children(ast):
                results.append(self.eval_volatiles_from_ast(ast, c, cell))
        return list(flatten(results, only_lists = True))

    def reduce(self, inputs = [], outputs = [], original_cells = None, original_nr = None):
        independent_volatiles, all_volatiles = self.detect_alive(inputs, outputs)
        new_cells, new_named_ranges, new_volatiles = self.fix_volatiles(independent_volatiles)

        for address,cell in self.cells.items():
            if cell.is_range:
                del self.cells[address]
        
        for address,new_cell in new_cells.items():
            # print "==============="
            # print address
            # print original_cells[address].formula
            # print new_cell.formula
            original_cells[address] = new_cell
        for address,new_cell in new_named_ranges.items():
            original_nr[address] = new_cell
        self.cells = original_cells
        self.named_ranges = original_nr
        self.volatiles = new_volatiles    
        self.Range = RangeFactory(self.cells)
        self.G = None
        self.gen_graph(inputs, outputs)
        # new_cells, G = graph_from_seeds(map(lambda x: self.cells[x], outputs), self)
        # print "Graph construction done, %s nodes, %s edges, %s new_cells entries" % (len(G.nodes()),len(G.edges()),len(new_cells))


    def detect_alive(self, inputs = [], outputs = []):

        volatile_arguments, all_volatiles = self.find_volatile_arguments(outputs)

        # for arg in volatile_arguments:
        #     if arg in self.cells and self.cells[arg].is_range:
        #         del self.cell[arg]

        independent_volatiles = all_volatiles.copy()

        # go down the tree and list all cells that are volatile arguments
        todo = [self.cells[input] for input in inputs]
        done = set()

        volatiles_concerned = set()

        while len(todo) > 0:
            cell = todo.pop()

            if cell not in done:
                if cell.address() in volatile_arguments.keys():
                    for vc in volatile_arguments[cell.address()]:
                        volatiles_concerned.add(vc)
                        try:
                            independent_volatiles.remove(vc)
                        except:
                            print "WARNING could not remove from independent_volatiles"
                            pass

                for child in self.G.successors_iter(cell):
                    todo.append(child)
  
                done.add(cell)
        print "Number of volatiles impacted by inputs: ", len(volatiles_concerned)

        return independent_volatiles, all_volatiles


    def find_volatile_arguments(self, outputs = None):

        # 1) gather all occurence of volatile 
        all_volatiles = set()

        # if outputs is None:
        # 1.1) from all cells
        for volatile in self.volatiles:
            cell = self.cells[volatile]
            if cell.formula:
                all_volatiles.add((cell.formula, cell.address(), cell.sheet if cell.sheet is not None else None))
            else:
                raise Exception('Volatiles should always have a formula')

        # else:
        #     # 1.2) from the outputs while climbing up the tree
        #     todo = [self.cells[output] for output in outputs]
        #     done = set()
        #     while len(todo) > 0:
        #         cell = todo.pop()

        #         if cell not in done:
        #             if cell.address() in self.volatiles:
        #                 if cell.formula:
        #                     all_volatiles.add((cell.formula, cell.address(), cell.sheet if cell.sheet is not None else None))
        #                 else:
        #                     raise Exception('Volatiles should always have a formula')

        #             for parent in self.G.predecessors_iter(cell): # climb up the tree      
        #                 todo.append(parent)

        #             done.add(cell)
        print "Total number of volatiles ", len(all_volatiles)

        # 2) extract the arguments from these volatiles
        done = set()
        volatile_arguments = {}

        for formula, address, sheet in all_volatiles:
            if formula not in done:
                if sheet:
                    parsed = parse_cell_address(address)
                else:
                    parsed = ""
                e = shunting_yard(formula, self.named_ranges, ref=parsed, tokenize_range = True)
                ast,root = build_ast(e)
                code = root.emit(ast)
                
                for a in list(flatten(self.get_volatile_arguments_from_ast(ast, root, sheet))):
                    if a in volatile_arguments:
                        volatile_arguments[a].append((formula, address, sheet))
                    else:
                        volatile_arguments[a] = [(formula, address, sheet)]

                done.add(formula) 

        return volatile_arguments, all_volatiles


    def get_arguments_from_ast(self, ast, node, sheet):
        arguments = []

        for c in node.children(ast):
            if c.tvalue == ":":
                arg_range =  reverse_rpn(c, ast)
                for elem in resolve_range(arg_range, False, sheet)[0]:
                    arguments += [elem]
            if c.ttype == "operand":
                if not is_number(c.tvalue):
                    if sheet is not None and "!" not in c.tvalue and c.tvalue not in self.named_ranges:
                        arguments += [sheet + "!" + c.tvalue]
                    else:
                        arguments += [c.tvalue]
            else:
                arguments += [self.get_arguments_from_ast(ast, c, sheet)]

        return arguments

    def get_volatile_arguments_from_ast(self, ast, node, sheet):
        arguments = []

        if node.token.tvalue in self.volatile_to_remove:
            for c in node.children(ast)[1:]:
                if c.ttype == "operand":
                    if not is_number(c.tvalue):
                        if sheet is not None and "!" not in c.tvalue and c.tvalue not in self.named_ranges:
                            arguments += [sheet + "!" + c.tvalue]
                        else:
                            arguments += [c.tvalue]
                else:
                        arguments += [self.get_arguments_from_ast(ast, c, sheet)]
        else:
            for c in node.children(ast):
                arguments += [self.get_volatile_arguments_from_ast(ast, c, sheet)]

        return arguments
      
    
    def set_value(self, address, val):

        self.reset_buffer = set()

        if address not in self.cells:
            raise Exception("Address not present in graph.")

        address = address.replace('$','')
        cell = self.cells[address]

        # when you set a value on cell, its should_eval flag is set to 'never' so its formula is not used until set free again => sp.activate_formula()
        self.fix_cell(address)

        # case where the address refers to a range
        if self.cells[address].is_range: 
            cells_to_set = []
            # for a in self.cells[address].range.addresses:
                # if a in self.cells:
                #     cells_to_set.append(self.cells[a])
                #     self.fix_cell(a)

            if type(val) != list:
                val = [val]*len(cells_to_set)

            self.reset(cell)
            cell.range.values = val

        # case where the address refers to a single value
        else:
            if address in self.named_ranges: # if the cell is a named range, we need to update and fix the reference cell
                ref_address = self.named_ranges[address]
                
                if ref_address in self.cells:
                    ref_cell = self.cells[ref_address]
                else:
                    ref_cell = Cell(ref_address, None, value = val, formula = None, is_range = False, is_named_range = False )
                    self.add_cell(ref_cell)

                # self.fix_cell(ref_address)
                ref_cell.value = val

            if cell.value != val:
                if cell.value is None:
                    cell.value = 'notNone' # hack to avoid the direct return in reset() when value is None
                # reset the node + its dependencies
                self.reset(cell)
                # set the value
                cell.value = val

        for volatile in self.volatiles: # reset all volatiles
            self.reset(self.cells[volatile])

    def reset(self, cell):
        addr = cell.address()
        if cell.value is None and addr not in self.named_ranges: return

        # update cells
        if cell.should_eval != 'never':
            if not cell.is_range:
                cell.value = None

            self.reset_buffer.add(cell)
            cell.need_update = True

        for child in self.G.successors_iter(cell):
            if child not in self.reset_buffer:
                self.reset(child)

    def fix_cell(self, address):
        if address in self.cells:
            if address not in self.fixed_cells:
                cell = self.cells[address]
                self.fixed_cells[address] = cell.should_eval
                cell.should_eval = 'never'
        else:
            raise Exception('Cell %s not in cellmap' % address)

    def free_cell(self, address = None):
        if address is None:
            for addr in self.fixed_cells:
                cell = self.cells[addr]

                cell.should_eval = 'always' # this is to be able to correctly reinitiliaze the value
                if cell.python_expression is not None:
                    self.eval_ref(addr)
                
                cell.should_eval = self.fixed_cells[addr]
            self.fixed_cells = {}

        elif address in self.cells:
            cell = self.cells[address]

            cell.should_eval = 'always' # this is to be able to correctly reinitiliaze the value
            if cell.python_expression is not None:
                self.eval_ref(address)
            
            cell.should_eval = self.fixed_cells[address]
            self.fixed_cells.pop(address, None)
        else:
            raise Exception('Cell %s not in cellmap' % address)

    def print_value_tree(self,addr,indent):
        cell = self.cells[addr]
        print "%s %s = %s" % (" "*indent,addr,cell.value)
        for c in self.G.predecessors_iter(cell):
            self.print_value_tree(c.address(), indent+1)

    def build_volatile(self, volatile):
        if not isinstance(volatile, RangeCore):
            vol_range = self.cells[volatile].range
        else:
            vol_range = volatile

        start = eval(vol_range.reference['start'])
        end = eval(vol_range.reference['end'])

        vol_range.build('%s:%s' % (start, end), debug = True)

    def build_volatiles(self):

        for volatile in self.volatiles:
            vol_range = self.cells[volatile].range

            start = eval(vol_range.reference['start'])
            end = eval(vol_range.reference['end'])

            vol_range.build('%s:%s' % (start, end), debug = True)

    def eval_ref(self, addr1, addr2 = None, ref = None):
        debug = False

        if isinstance(addr1, ExcelError):
            return addr1
        elif isinstance(addr2, ExcelError):
            return addr2
        else:
            if addr1 in self.cells:
                cell1 = self.cells[addr1]
            elif addr2 is None:
                if self.debug:
                    print 'WARNING in eval_ref: address %s not found in cellmap, returning #NULL' % addr1
                return ExcelError('#NULL', 'Cell %s is empty' % addr1)
            if addr2 == None:
                if cell1.is_range:

                    if cell1.range.is_volatile:
                        self.build_volatile(cell1.range)

                    associated_addr = RangeCore.find_associated_cell(ref, cell1.range)

                    if associated_addr: # if range is associated to ref, no need to return/update all range
                        return self.evaluate(associated_addr)
                    else:
                        range_name = cell1.address()
                        if cell1.need_update:
                            self.update_range(cell1.range)
                            range_need_update = True
                            
                            for c in self.G.successors_iter(cell1): # if a parent doesnt need update, then cell1 doesnt need update
                                if not c.need_update:
                                    range_need_update = False
                                    break

                            cell1.need_update = range_need_update
                            return cell1.range
                        else:
                            return cell1.range

                elif addr1 in self.named_ranges or not is_range(addr1):
                    val = self.evaluate(addr1)
                    return val
                else: # addr1 = Sheet1!A1:A2 or Sheet1!A1:Sheet1!A2
                    addr1, addr2 = addr1.split(':')
                    if '!' in addr1:
                        sheet = addr1.split('!')[0]
                    else:
                        sheet = None
                    if '!' in addr2:
                        addr2 = addr2.split('!')[1]

                    temp_range = self.Range('%s:%s' % (addr1, addr2))
                    self.update_range(temp_range)
                    return temp_range
            else:  # addr1 = Sheet1!A1, addr2 = Sheet1!A2
                if '!' in addr1:
                    sheet = addr1.split('!')[0]
                else:
                    sheet = None
                if '!' in addr2:
                    addr2 = addr2.split('!')[1]
                temp_range = self.Range('%s:%s' % (addr1, addr2))
                self.update_range(temp_range)
                return temp_range

    def update_range(self, range):
        # This function loops through its Cell references to evaluate the ones that need so
        # This uses Spreadsheet.pending dictionary, that holds the addresses of the Cells that are being calculated
        
        debug = False

        for index, key in enumerate(range.order):
            addr = get_cell_address(range.sheet, key)
            a = self.cells[addr]
            if self.cells[addr].need_update:
                self.evaluate(addr)
            

    def evaluate(self,cell,is_addr=True):
        if is_addr:
            try:
                cell = self.cells[cell]
            except:
                if self.debug:
                    print 'WARNING: Empty cell at ' + cell
                return ExcelError('#NULL', 'Cell %s is empty' % cell)    

        # no formula, fixed value
        if cell.should_eval == 'normal' and not cell.need_update and cell.value is not None or not cell.formula or cell.should_eval == 'never':
            return cell.value
        try:
            if cell.is_range:
                for child in cell.range.cells:
                    self.evaluate(child.address())
            elif cell.compiled_expression != None:
                vv = eval(cell.compiled_expression)
                if isinstance(vv, RangeCore): # this should mean that vv is the result of RangeCore.apply_all, but with only one value inside
                    cell.value = vv.values[0]
                else:
                    cell.value = vv
            else:
                cell.value = 0
            
            cell.need_update = False
            
            # DEBUG: saving differences
            if self.save_history:

                def is_almost_equal(a, b, precision = 0.001):
                    if is_number(a) and is_number(b):
                        return abs(float(a) - float(b)) <= precision
                    elif (a is None or a == 'None') and (b is None or b == 'None'):
                        return True
                    else:
                        return a == b

                if cell.address() in self.history:
                    ori_value = self.history[cell.address()]['original']
                    
                    if 'new' not in self.history[cell.address()].keys():
                        if type(ori_value) == list and type(cell.value) == list \
                                and all(map(lambda (x, y): not is_almost_equal(x, y), zip(ori_value, cell.value))) \
                            or not is_almost_equal(ori_value, cell.value):

                            self.count += 1
                            self.history[cell.address()]['formula'] = str(cell.formula)
                            self.history[cell.address()]['priority'] = self.count
                            self.history[cell.address()]['python'] = str(cell.python_expression)

                            if self.count == 1:
                                self.history['ROOT_DIFF'] = self.history[cell.address()]
                                self.history['ROOT_DIFF']['cell'] = cell.address()

                    self.history[cell.address()]['new'] = cell.value
                else:
                    self.history[cell.address()] = {'new': cell.value}

        except Exception as e:
            if e.message is not None and e.message.startswith("Problem evalling"):
                raise e
            else:
                raise Exception("Problem evalling: %s for %s, %s" % (e,cell.address(),cell.python_expression)) 

        return cell.value


    def gen_graph(self, inputs = [], outputs = []):
        print '___### Generating Graph ###___', len(self.cells),len(self.named_ranges), len(self.volatiles)
        # for address,cell in self.cells.items():
        #     print address, cell.formula

        if len(outputs) == 0:
            preseeds = set(list(flatten(self.cells.keys())) + self.named_ranges.keys()) # to have unicity
        else:
            preseeds = set(outputs)
        
        preseeds = list(preseeds) # to be able to modify the list

        seeds = []
        for o in preseeds:
            if o in self.named_ranges:
                reference = self.named_ranges[o]

                if is_range(reference):
                    if 'OFFSET' in reference or 'INDEX' in reference:
                        start_end = prepare_volatile(reference, self.named_ranges)
                        rng = self.Range(start_end)
                        self.volatiles.add(o)
                    else:
                        rng = self.Range(reference)

                    for address in rng.addresses:
                        preseeds.append(address)
                    virtual_cell = Cell(o, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                    seeds.append(virtual_cell)
                else:
                    # might need to be changed to actual self.cells Cell, not a copy
                    if 'OFFSET' in reference or 'INDEX' in reference:
                        self.volatiles.add(o)

                    value = self.cells[reference].value if reference in self.cells else None
                    virtual_cell = Cell(o, None, value = value, formula = reference, is_range = False, is_named_range = True)
                    seeds.append(virtual_cell)
            else:
                if is_range(o):
                    rng = self.Range(o)
                    for address in rng.addresses: # this is avoid pruning deletion
                        preseeds.append(address)
                    virtual_cell = Cell(o, None, value = rng, formula = o, is_range = True, is_named_range = True )
                    seeds.append(virtual_cell)
                else:
                    seeds.append(self.cells[o])

        seeds = set(seeds)
        print "Seeds %s cells" % len(seeds)
        outputs = set(preseeds) if len(outputs) > 0 else [] # seeds and outputs are the same when you don't specify outputs

        new_cells, G = graph_from_seeds(seeds, self)

        if len(inputs) != 0: # otherwise, we'll set inputs to new_cells inside Spreadsheet
            inputs = list(set(inputs))

            # add inputs that are outside of calculation chain
            for i in inputs:
                if i not in new_cells:
                    if i in self.named_ranges:
                        reference = self.named_ranges[i]
                        if is_range(reference):

                            rng = self.Range(reference)
                            for address in rng.addresses: # this is avoid pruning deletion
                                inputs.append(address)
                            virtual_cell = Cell(i, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                            new_cells[i] = virtual_cell
                            G.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain

                        else:
                            # might need to be changed to actual self.cells Cell, not a copy
                            virtual_cell = Cell(i, None, value = self.cells[reference].value, formula = reference, is_range = False, is_named_range = True)
                            new_cells[i] = virtual_cell
                            G.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain
                    else:
                        if is_range(i):
                            rng = self.Range(i)
                            for address in rng.addresses: # this is avoid pruning deletion
                                inputs.append(address)
                            virtual_cell = Cell(i, None, value = rng, formula = o, is_range = True, is_named_range = True )
                            new_cells[i] = virtual_cell
                            G.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain
                        else:
                            new_cells[i] = self.cells[i]
                            G.add_node(self.cells[i]) # edges are not needed here since the input here is not in the calculation chain

            inputs = set(inputs)

        self.G = G
        self.cells = new_cells

        print "Graph construction done, %s nodes, %s edges, %s new_cells entries" % (len(G.nodes()),len(G.edges()),len(new_cells))

