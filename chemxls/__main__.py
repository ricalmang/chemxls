import os, functools, math, sys, itertools, re, time, threading, xlwt, configparser
import numpy as np
import tkinter as tk
from xlwt import Workbook
from tkinter import ttk, filedialog, messagebox

#CONSTANTS
# Data from: Pyykko, P. and Atsumi, M., Chem. Eur. J. 2009, 15, 186.
element_radii=[
    ["None",None],['H'   ,  32],['He'  ,  46],['Li'  , 133],['Be'  , 102],['B'   ,  85],['C'   ,  75],
    ['N'   ,  71],['O'   ,  63],['F'   ,  64],['Ne'  ,  67],['Na'  , 155],['Mg'  , 139],['Al'  , 126],
	['Si'  , 116],['P'   , 111],['S'   , 103],['Cl'  ,  99],['Ar'  ,  96],['K'   , 196],['Ca'  , 171],
	['Sc'  , 148],['Ti'  , 136],['V'   , 134],['Cr'  , 122],['Mn'  , 119],['Fe'  , 116],['Co'  , 111],
	['Ni'  , 110],['Cu'  , 112],['Zn'  , 118],['Ga'  , 124],['Ge'  , 121],['As'  , 121],['Se'  , 116],
	['Br'  , 114],['Kr'  , 117],['Rb'  , 210],['Sr'  , 185],['Y'   , 163],['Zr'  , 154],['Nb'  , 147],
	['Mo'  , 138],['Tc'  , 128],['Ru'  , 125],['Rh'  , 125],['Pd'  , 120],['Ag'  , 128],['Cd'  , 136],
    ['In'  , 142],['Sn'  , 140],['Sb'  , 140],['Te'  , 136],['I'   , 133],['Xe'  , 131],['Cs'  , 232],
	['Ba'  , 196],['La'  , 180],['Ce'  , 163],['Pr'  , 176],['Nd'  , 174],['Pm'  , 173],['Sm'  , 172],
	['Eu'  , 168],['Gd'  , 169],['Tb'  , 168],['Dy'  , 167],['Ho'  , 166],['Er'  , 165],['Tm'  , 164],
	['Yb'  , 170],['Lu'  , 162],['Hf'  , 152],['Ta'  , 146],['W'   , 137],['Re'  , 131],['Os'  , 129],
	['Ir'  , 122],['Pt'  , 123],['Au'  , 124],['Hg'  , 133],['Tl'  , 144],['Pb'  , 144],['Bi'  , 151],
	['Po'  , 145],['At'  , 147],['Rn'  , 142],['Fr'  , 223],['Ra'  , 201],['Ac'  , 186],['Th'  , 175],
    ['Pa'  , 169],['U'   , 170],['Np'  , 171],['Pu'  , 172],['Am'  , 166],['Cm'  , 166],['Bk'  , 168],
	['Cf'  , 168],['Es'  , 165],['Fm'  , 167],['Md'  , 173],['No'  , 176],['Lr'  , 161],['Rf'  , 157],
	['Db'  , 149],['Sg'  , 143],['Bh'  , 141],['Hs'  , 134],['Mt'  , 129],['Ds'  , 128],['Rg'  , 121],
	['Cn'  , 122],['Nh'  , 136],['Fl'  , 143],['Mc'  , 162],['Lv'  , 175],['Ts'  , 165],['Og'  , 157]]
elements = tuple(i[0] for i in element_radii)
keywords = \
    ['1-bromo-2-methylpropane', '1-bromooctane', '1-bromopentane', '1-bromopropane', '1-butanol',
    '1-chlorohexane', '1-chloropentane', '1-chloropropane', '1-decanol', '1-fluorooctane', '1-heptanol',
    '1-hexanol', '1-hexene', '1-hexyne', '1-iodobutane', '1-iodohexadecane', '1-iodopentane',
    '1-iodopropane', '1-nitropropane', '1-nonanol', '1-pentanol', '1-pentene', '1-propanol',
    '1-trichloroethane', '2-bromopropane', '2-butanol', '2-chlorobutane', '2-dibromoethane',
    '2-dichloroethene', '2-dimethylcyclohexane', '2-ethanediol', '2-heptanone', '2-hexanone',
    '2-methoxyethanol', '2-methyl-1-propanol', '2-methyl-2-propanol', '2-methylpentane',
    '2-methylpyridine', '2-nitropropane', '2-octanone', '2-pentanone', '2-propanol', '2-propen-1-ol',
    '2-trichloroethane', '2-trifluoroethanol', '3-methylpyridine', '3-pentanone', '4-dimethylpentane',
    '4-dimethylpyridine', '4-dioxane', '4-heptanone', '4-methyl-2-pentanone', '4-methylpyridine',
    '4-trimethylbenzene', '4-trimethylpentane', '5-nonanone', '6-dimethylpyridine', 'a-chlorotoluene',
    'aceticacid', 'acetone', 'acetonitrile', 'acetophenone', 'allcheck', 'aniline', 'anisole', 'apfd',
    'argon', 'b1b95', 'b1lyp', 'b3lyp', 'b3p86', 'b3pw91', 'b971', 'b972', 'b97d', 'b97d3', 'benzaldehyde',
    'benzene', 'benzonitrile', 'benzylalcohol', 'betanatural', 'bhandh', 'bhandhlyp', 'bromobenzene',
    'bromoethane', 'bromoform', 'butanal', 'butanoicacid', 'butanone', 'butanonitrile', 'butylamine',
    'butylethanoate', 'calcall', 'calcfc', 'cam-b3lyp', 'carbondisulfide', 'carbontetrachloride',
    'cartesian', 'checkpoint', 'chkbasis', 'chlorobenzene', 'chloroform', 'cis-1', 'cis-decalin',
    'connectivity', 'counterpoise', 'cyclohexane', 'cyclohexanone', 'cyclopentane', 'cyclopentanol',
    'cyclopentanone', 'd95v', 'decalin-mixture', 'def2qzv', 'def2qzvp', 'def2qzvpp', 'def2sv', 'def2svp',
    'def2svpp', 'def2tzv', 'def2tzvp', 'def2tzvpp', 'density', 'densityfit', 'dibromomethane',
    'dibutylether', 'dichloroethane', 'dichloromethane', 'diethylamine', 'diethylether', 'diethylsulfide',
    'diiodomethane', 'diisopropylether', 'dimethyldisulfide', 'dimethylsulfoxide', 'diphenylether',
    'dipropylamine', 'e-2-pentene', 'empiricaldispersion', 'ethanethiol', 'ethanol', 'ethylbenzene',
    'ethylethanoate', 'ethylmethanoate', 'ethylphenylether', 'extrabasis', 'extradensitybasis', 'finegrid',
    'fluorobenzene', 'formamide', 'formicacid', 'freq', 'full', 'gd3bj', 'genecp', 'geom', 'gfinput',
    'gfprint', 'hcth', 'hcth147', 'hcth407', 'hcth93', 'heptane', 'hexanoicacid', 'hissbpbe', 'hseh1pbe',
    'integral', 'iodobenzene', 'iodoethane', 'iodomethane', 'isopropylbenzene', 'isoquinoline', 'kcis',
    'krypton', 'lanl2dz', 'lanl2mb', 'lc-wpbe', 'loose', 'm-cresol', 'm-xylene', 'm062x', 'm06hf', 'm06l',
    'm11l', 'maxcycles', 'maxstep', 'mesitylene', 'methanol', 'methylbenzoate', 'methylbutanoate',
    'methylcyclohexane', 'methylethanoate', 'methylmethanoate', 'methylpropanoate', 'minimal', 'mn12l',
    'mn12sx', 'modredundant', 'mpw1lyp', 'mpw1pbe', 'mpw1pw91', 'mpw3pbe', 'n-butylbenzene', 'n-decane',
    'n-dimethylacetamide', 'n-dimethylformamide', 'n-dodecane', 'n-hexadecane', 'n-hexane',
    'n-methylaniline', 'n-methylformamide-mixture', 'n-nonane', 'n-octane', 'n-octanol', 'n-pentadecane',
    'n-pentane', 'n-undecane', 'n12sx', 'nitrobenzene', 'nitroethane', 'nitromethane', 'noeigentest',
    'nofreeze', 'noraman', 'nosymm', 'nprocshared', 'o-chlorotoluene', 'o-cresol', 'o-dichlorobenzene',
    'o-nitrotoluene', 'o-xylene', 'o3lyp', 'ohse1pbe', 'ohse2pbe', 'oniom', 'output', 'p-isopropyltoluene',
    'p-xylene', 'pbe1pbe', 'pbeh', 'pbeh1pbe', 'pentanal', 'pentanoicacid', 'pentylamine',
    'pentylethanoate', 'perfluorobenzene', 'pkzb', 'population', 'propanal', 'propanoicacid',
    'propanonitrile', 'propylamine', 'propylethanoate', 'pseudo', 'pw91', 'pyridine', 'qst2', 'qst3',
    'quinoline', 'qzvp', 'rdopt', 'read', 'readfc', 'readfreeze', 'readopt', 'readoptimize', 'regular',
    'restart', 's-dioxide', 'savemixed', 'savemulliken', 'savenbos', 'savenlmos', 'scrf', 'sddall',
    'sec-butylbenzene', 'sogga11', 'sogga11x', 'solvent', 'spinnatural', 'tert-butylbenzene',
    'tetrachloroethene', 'tetrahydrofuran', 'tetrahydrothiophene-s', 'tetralin', 'thcth', 'thcthhyb',
    'thiophene', 'thiophenol', 'tight', 'toluene', 'tpss', 'tpssh', 'trans-decalin', 'tributylphosphate',
    'trichloroethene', 'triethylamine', 'tzvp', 'ultrafine', 'uncharged', 'v5lyp', 'verytight', 'vp86',
    'vsxc', 'vwn5', 'water', 'wb97', 'wb97x', 'wb97xd', 'wpbeh', 'x3lyp', 'xalpha', 'xenon',
    'xylene-mixture']
#GENERAL PURPOSE FUNCTIONS
def is_str_float(i):
	"""Check if a string can be converted into a float"""
	try: float(i); return True
	except ValueError: return False
	except TypeError: return False
def trim_str(string, max_len=40):
	assert type(string) == str
	assert type(max_len) == int
	if len(string) > max_len: return "..." + string[-max_len:]
	else: return string
def read_item(file_name):
	"""Reads an .xyz, .gjf, .com or .log item and returns a list of its contents ready for class instantiation"""
	with open(file_name,"r") as in_file:
		in_content = [file_name]
		in_content.extend(list(in_file.read().splitlines()))
	return in_content
def lock_release(func):
	def new_func(*args, **kw):
		global frame_a, frame_b
		if frame_a.lock or frame_b.lock: return None
		frame_a.lock, frame_b.lock = True, True
		for a in frame_a.check_buttons: a.config(state=tk.DISABLED)
		for a in frame_b.check_buttons: a.config(state=tk.DISABLED)
		for a in frame_a.buttons: a.config(state=tk.DISABLED)
		for a in frame_b.buttons: a.config(state=tk.DISABLED)
		result = func(*args, **kw)
		for a in frame_a.check_buttons: a.config(state=tk.NORMAL)
		for a in frame_b.check_buttons: a.config(state=tk.NORMAL)
		for a in frame_a.buttons: a.config(state=tk.NORMAL)
		for a in frame_b.buttons: a.config(state=tk.NORMAL)
		frame_a.lock, frame_b.lock = False, False
		return result
	return new_func
#DATA FILE CLASSES
class LogFile:
	calc_types = ["TS","Red","IRC","Opt","SP"]
	def __init__(self,file_content,fragment_link_one=False):
		self.list = file_content
		self.lenght = len(self.list)
		self.name = self.list[0].strip()
		self.empty_line_idxs = []
		self.charge_mult = None
		self.input_geom_idx = None
		self.start_xyz_idxs = []
		self.end_resume_idxs = []
		self.start_resume_idxs = []
		self.linked_job_idxs = []
		self.multi_dash_idxs =[]
		self.scf_done = []
		####.thermal = ["ZPC","TCE","TCH","TCG","SZPE","STE","STH","STG"]
		self.thermal = [None, None, None, None, None , None, None, None]
		self.oc_orb_energies = []
		self.uno_orb_energies = []
		self.hash_line_idxs = []
		self.norm_term_idxs = []
		self.errors = []
		self.irc_points = []
		self.scan_points = []
		self.opt_points = []
		self.force_const_mat = []
		self.distance_matrix = []
		self.s_squared = []
		self.muliken_spin_densities_idxs = []
		self.muliken_charge_idxs = []
		self.chelpg_charge_idxs = []
		self.pop_analysis_idxs = []
		self.npa_start_idxs = []
		self.npa_end_idxs = []
		self.apt_charge_idxs =[]
		for i,a in enumerate(a.strip() for a in self.list):
			# i = index
			# a = line.strip()
			# b = line.split()
			# c = len(b)
			if a == "":                                                         self.empty_line_idxs.append(i); continue
			if a[-1] == "@":                                                    self.end_resume_idxs.append(i); continue
			elif a[0] == "1":
				if a.startswith(r"1\1"):                                      self.start_resume_idxs.append(i); continue
			if a[0].isdigit() or a[0].islower():                                                                continue
			elif a[0] == "-":
				if a.startswith("------"):                                      self.multi_dash_idxs.append(i); continue
			elif a[0] == "!":
				b = a.split(); c = len(b)
				if c == 4:
					condition_a = all(x in y for x,y in zip(b,("!",["Optimized","Non-Optimized"],"Parameters","!")))
					if condition_a:                                          self.scan_points.append([i,b[1]]); continue
			elif a[0] == "A":
				text_a = "Alpha  occ. eigenvalues --"
				text_b = "Alpha virt. eigenvalues --"
				text_c = "Atom  No          Natural Electron Configuration"
				text_d = "APT charges:"
				if a.startswith(text_a):                                        self.oc_orb_energies.append(i); continue
				elif a.startswith(text_b):                                     self.uno_orb_energies.append(i); continue
				elif a.split() == text_c.split():                                  self.npa_end_idxs.append(i); continue
				elif a.startswith(text_d):                                      self.apt_charge_idxs.append(i); continue

			elif a[0] == "C":
				b = a.split(); c = len(b)
				if all((a.startswith("Charge"),self.charge_mult is None, c == 6)):
					pattern = ("Charge", "=", "Multiplicity", "=")
					if all(x == b[n] for x,n in zip(pattern,(0,1,3,4))):
						self.input_geom_idx = i;                                    self.charge_mult = b[2::3]; continue
			elif a[0] == "D":
				if a.startswith("Distance matrix (angstroms):"):   				self.distance_matrix.append(i); continue
			elif a[0] == "E":
				if a.startswith("Error"):                                                self.errors.append(i); continue
				elif a.startswith("ESP charges:"):                           self.chelpg_charge_idxs.append(i);continue
			elif a[0] == "F":
				if a.startswith("Full mass-weighted force constant matrix:"):   self.force_const_mat.append(i); continue
			elif a[0] == "I":
				if a == "Input orientation:":                                self.start_xyz_idxs.append(i + 5); continue
			elif a[0] == "L":
				if a.startswith("Link1:"):
					self.linked_job_idxs.append(i)
					if fragment_link_one:
						self.lenght = len(self.list[:i])
						self.link_one = [self.list[0]]
						try:
							_ = self.list[i+1]
							self.link_one.extend(self.list[i + 1:])
							self.link_one = LogFile(self.link_one,fragment_link_one)
							self.list = self.list[:i]
							break
						except IndexError:
							pass
			elif a[0] == "N":
				if a.startswith("Normal termination of Gaussian"):               self.norm_term_idxs.append(i); continue
			elif a[0] == "M":
				if a.startswith("Mulliken charges and spin densities:"):
					pass;                                           self.muliken_spin_densities_idxs.append(i); continue
				elif a.startswith("Mulliken charges:"):
					pass;                                                    self.muliken_charge_idxs.append(i);continue
				elif a.startswith("Molecular Orbital Coefficients:"):
					pass;                                                      self.pop_analysis_idxs.append(i);continue
			elif a[0] == "P":
				b = a.split(); c = len(b)
				if c != 6 or any(x != b[n] for x,n in zip(["Point","Number:","Path","Number:"],[0,1,3,4])):     continue
				if any(not b[n].isnumeric() for n in [2, 5]):                                                   continue
				else:                                                  self.irc_points.append([i, b[5], b[2]]); continue
			elif a[0] == "S":
				b = a.split(); c = len(b)
				if a == "Standard orientation:":                             self.start_xyz_idxs.append(i + 5); continue
				elif a.startswith("SCF Done:") and c > 5:                       self.scf_done.append([i,b[4]]); continue
				elif a.startswith("S**2 before annihil"):self.s_squared.append([i,b[3].replace(",",""),b[-1]]); continue
				elif a.startswith("Sum of electronic and zero-point Energies="):       self.thermal[4] = b[-1]; continue
				elif a.startswith("Sum of electronic and thermal Energies="):          self.thermal[5] = b[-1]; continue
				elif a.startswith("Sum of electronic and thermal Enthalpies="):        self.thermal[6] = b[-1]; continue
				elif a.startswith("Sum of electronic and thermal Free Energies="):     self.thermal[7] = b[-1]; continue
				elif a.startswith("Step") and c == 9:
					x = ["Step", "number", "out", "of", "a", "maximum", "of"]
					y = [0, 1, 3, 4, 5, 6, 7]
					z = all(b[n].isnumeric() for n in [2, 8])
					if all(d == b[n] for d,n in zip(x,y)) and z:                     self.opt_points.append(i); continue
			elif a[0] == "T":
				b = a.split()
				if a.startswith("Thermal correction to Energy="):                      self.thermal[1] = b[-1]; continue
				elif a.startswith("Thermal correction to Enthalpy="):                  self.thermal[2] = b[-1]; continue
				elif a.startswith("Thermal correction to Gibbs Free Energy="):         self.thermal[3] = b[-1]; continue
			elif a[0] == "Z":
				b = a.split()
				if a.startswith("Zero-point correction="):                             self.thermal[0] = b[-2]; continue
			elif a[0] == "#":                                                    self.hash_line_idxs.append(i); continue
			elif a[0] == "*":
				if a.replace("*","").startswith("Gaussian NBO Version 3.1"):
					pass;                                                         self.npa_start_idxs.append(i);continue
		#--------------------------------------------POST PROCESSING----------------------------------------------------
		x = None if self.start_xyz_idxs is None else [min(a for a in self.multi_dash_idxs if a > b) for b in self.start_xyz_idxs]
		self.end_xyz_idxs = x
		self.scan_end = [min(a for a in self.multi_dash_idxs if a > b[0]) for b in self.scan_points]
		try:
			x = [self.list[b:min(a for a in self.empty_line_idxs if a > b)] for b in self.force_const_mat]
			self.displ_block = x
		except Exception as e:
			print("Error while finding vibrational frequencies of log file")
			print(e)
			print(self.name)
			self.displ_block = []
		# --------------------------------------------------ASSURANCE---------------------------------------------------
		self.init_errors = []
		#if self.charge_mult is None:
		#	self.init_errors.append("Charge and multiplicity could not be identified!")
		#if len(self.start_resume_idxs) != len(self.end_resume_idxs):
		#	self.init_errors.append("Inconsistent resumes")
		#if len(self.name.split()) != 1:
		#	self.init_errors.append("Name should not contain empty spaces or be empty")
		#if not self.list[1].strip().startswith("Entering Gaussian System"):
		#	self.init_errors.append("Is this a Gaussian log file?")
		#if not self.start_xyz_idxs is None:
		#	if len(self.start_xyz_idxs) != len(self.end_xyz_idxs):
		#		self.init_errors.append("Found an inconsistent number of geometries")
		#if not any([self.homo is None, self.lumo is None]):
		#	if self.homo > self.lumo:
		#		self.init_errors.append("Lumo is lower than homo?")
		#if self.init_errors:
		#	for a in self.init_errors: print(a)
		#	print("Errors above were found on file\n{}".format(self.name))
	@functools.lru_cache(maxsize=1)
	def loghelp(self):
		for a in vars(self):
			if a != "list":
				print(a.upper(),"--->",getattr(self,a))
	@functools.lru_cache(maxsize=1)
	def xyz_cord_block(self,start_idx,end_idx):
		data = [a.split() for a in self.list[start_idx:end_idx]]
		return [[elements[int(l[1])],*[l[i] for i in [3,4,5]]] for l in data]
	@functools.lru_cache(maxsize=1)
	def last_cord_block(self):
		if not all([self.xyz_cord_block, self.end_xyz_idxs]):
			if self.last_log_abstract:
				print("WARNING: Coordinates will be taken from the last job abstract:")
				print("lines {} - {} of file:".format(self.start_resume_idxs[-1],self.end_resume_idxs[-1]))
				print("{}".format(self.name))
				return self.last_log_abstract.xyz_object().cord_block()
			else: return None
		else:
			return self.xyz_cord_block(self.start_xyz_idxs[-1],self.end_xyz_idxs[-1])
	@functools.lru_cache(maxsize=1)
	def first_cord_block(self):
		if not all([self.start_xyz_idxs,self.end_xyz_idxs]):
			if self.input_geom_idx:
				coordinates = []
				for i,a in enumerate(self.list[self.input_geom_idx:]):
					if i > 5 and not coordinates: break
					a = a.split()
					if len(a) == 4:
						if a[0] in elements and all(is_str_float(a[n]) for n in [1, 2, 3]):
							coordinates.append(a)
						elif coordinates: break
					elif coordinates: break
				return coordinates
			else: return None
		else:
			return self.xyz_cord_block(self.start_xyz_idxs[0],self.end_xyz_idxs[0])
	@functools.lru_cache(maxsize=1)
	def _n_atoms(self):
		if self.last_cord_block():
			return len(self.last_cord_block())
		elif self.first_cord_block():
			return len(self.first_cord_block())

	def any_xyz_obj(self,a_idx,b_idx,title=" ",name=False):
		if name == False: name = self.name
		return XyzFile([name, self.n_atoms, title, *(" ".join(l) for l in self.xyz_cord_block(a_idx,b_idx))])
	@functools.lru_cache(maxsize=1)
	def last_xyz_obj(self):
		if self.last_cord_block():
			return XyzFile([self.name,self.n_atoms," ",*(" ".join(l) for l in self.last_cord_block())])
		else:
			return None
	@functools.lru_cache(maxsize=1)
	def first_xyz_obj(self):
		return XyzFile([self.name,self.n_atoms," ",*(" ".join(l) for l in self.first_cord_block())])
	@functools.lru_cache(maxsize=1)
	def low_e_xyz_obj(self):
		if self.calc_type == "SP": return None
		else:
			xyzs = {"TS":self.opt,"Red":self.scan_geoms,"IRC":self.irc,"Opt":self.opt}[self.calc_type]()
			if len(xyzs) == 0: return None
			else: return sorted(xyzs,key= lambda x: float(x.title()) if is_str_float(x.title()) else 1)[0]
	@functools.lru_cache(maxsize=1)
	def _calc_type(self):
		if self.raw_route:
			r_sect = self.raw_route_keys
			if any(a in r_sect for a in ("ts", "qst2","qst3")): return "TS"
			elif any(True for a in r_sect if a in ("modredundant", "readoptimize", "readfreeze")): return "Red"
			elif "irc" in r_sect: return "IRC"
			elif "opt" in r_sect: return "Opt"
			else: return "SP"
		else: return "No data"
	@functools.lru_cache(maxsize=1)
	def _normal_termin(self):
		return any(True if "Normal termination of Gaussian" in l else False for l in self.list[-5:])
	def _error_msg(self):
		error_idxs = [a for a in self.errors if a + 5 > self.lenght]
		for n in [-4,-3,-2,-1]:
			if self.list[n].strip().startswith("galloc:  could not allocate memory."):
				error_idxs.append(n)
		if error_idxs: return " | ".join([self.list[n] for n in error_idxs])
		else: return "No data"
	@functools.lru_cache(maxsize=1)
	def needs_ref(self):
		if self.calc_type == "Opt" and self.last_freq:
			if self.last_freq.n_ifreq() == "0": return "No"
			else: return "Yes"
		elif self.calc_type == "TS" and self.last_freq:
			if self.last_freq.n_ifreq() == "1": return "No"
			else: return "Yes"
		else: return "-"
	@functools.lru_cache(maxsize=1)
	def irc(self):
		if not all([self.start_xyz_idxs,self.end_xyz_idxs,self.irc_points,self.scf_done]): return []
		points = self.irc_points
		scf = [max(self.scf_done,key=lambda x: x[0] if x[0] < a[0] else 0)[1] for a in points]
		a_idx = [max(self.start_xyz_idxs,key=lambda x: x if x < a[0] else 0) for a in points]
		b_idx = [max(self.end_xyz_idxs,key=lambda x: x if x < a[0] else 0) for a in points]
		points = [[*d[1:],c,self.any_xyz_obj(a,b,title=c)] for a,b,c,d in zip(a_idx,b_idx,scf,points)]
		path_a = sorted([a for a in points if a[0] == "1"], key = lambda x: int(x[1]), reverse=True)
		path_b = [a for a in points if a[0] == "2"]
		return [a[3] for a in [*path_a,*path_b]]
	@functools.lru_cache(maxsize=1)
	def opt(self):
		if not all([self.start_xyz_idxs,self.end_xyz_idxs,self.opt_points,self.scf_done]): return []
		points = self.opt_points
		scf = [max(self.scf_done,key=lambda x: x[0] if x[0] < a else 0)[1] for a in points]
		a_idx = [max(self.start_xyz_idxs,key=lambda x: x if x < a else 0) for a in points]
		b_idx = [max(self.end_xyz_idxs,key=lambda x: x if x < a else 0) for a in points]
		return [self.any_xyz_obj(a,b,title=c) for a,b,c in zip(a_idx,b_idx,scf)]
	@functools.lru_cache(maxsize=1)
	def scan_geoms(self):
		if not all([self.start_xyz_idxs, self.end_xyz_idxs, self.scan_points, self.scf_done]): return []
		geoms = []
		all_points = self.scan_points
		points = [a for a in all_points if a[0] < self.start_xyz_idxs[-1] and a[0] < self.end_xyz_idxs[-1]]
		points_removed = len(all_points) - len(points)
		if points_removed != 0:
			print(f"WARNING: {points_removed} Scan  points have been removed due to inconsistent number o geometries found")
		start_idx = [min(i for i in self.start_xyz_idxs if i > b[0])  for b in points]
		end_idx = [min(i for i in self.end_xyz_idxs if i > b[0]) for b in points]
		scf_idx = [max(i for i in self.scf_done if i[0] < b[0]) for b in points]
		for i,(a,b,c,d) in enumerate(zip(start_idx,end_idx,scf_idx,points)):
			name = self.name.replace(".log","_" + str(i+1)+".xyz")
			if d[1] == "Optimized": print("Optimized geometry found at line {}!".format(d[0]))
			elif d[1] == "Non-Optimized": print("Non-Optimized1 geometry found at line {}!".format(d[0]))
			geoms.append(self.any_xyz_obj(a,b,title=str(c[1]), name=name))
		if len(geoms) == 0:
			print("No Optimized geometries found for {} file".format(self.name()))
		return geoms
	@functools.lru_cache(maxsize=1)
	def _last_freq(self):
		return LogFreq(self.displ_block[-1]) if self.displ_block else False
	@functools.lru_cache(maxsize=1)
	def _last_log_abstract(self):
		if all([self.start_resume_idxs,self.end_resume_idxs]):
			x = ["".join([x.strip() for x in self.list[a:b]]).split("\\") for a,b in zip(self.start_resume_idxs,self.end_resume_idxs)]
			return LogAbstract(x[-1]) if x else None
	@functools.lru_cache(maxsize=1)
	def _xyz_from_dist_matrix(self):
		end_idx = lambda x: next(i for i,a in enumerate(self.list[x+1:],start=x+1) if not a.split()[0].isdigit())
		return [DistMatrix(self.list[a+1:end_idx(a)]) for a in self.distance_matrix]
	@functools.lru_cache(maxsize=1)
	def _last_muliken_spin_density(self):
		if self.muliken_spin_densities_idxs:
			end_idx = lambda x: next(i for i, a in enumerate(self.list[x + 1:], start=x + 1) if not a.split()[0].isdigit())
			return "\n".join(self.list[self.muliken_spin_densities_idxs[-1]:end_idx(self.muliken_spin_densities_idxs[-1]+1)])
	@functools.lru_cache(maxsize=1)
	def _last_internal_coord(self):
		if self.scan_points:
			end_idx = lambda x: next(i for i, a in enumerate(self.list[x + 1:], start=x + 1) if not a.strip().startswith("!"))
			return "\n".join(self.list[self.scan_points[-1][0]-1:end_idx(self.scan_points[-1][0]+5)+1])
	@functools.lru_cache(maxsize=1)
	def _last_muliken_charges(self):
		if self.muliken_charge_idxs:
			end_idx = lambda x: next(i for i, a in enumerate(self.list[x + 1:], start=x + 1) if not a.split()[0].isdigit())
			return "\n".join(self.list[self.muliken_charge_idxs[-1]:end_idx(self.muliken_charge_idxs[-1]+1)])
	@functools.lru_cache(maxsize=1)
	def _last_chelpg_charges(self):
		if self.chelpg_charge_idxs:
			end_idx = lambda x: next(i for i, a in enumerate(self.list[x + 1:], start=x + 1) if not a.split()[0].isdigit())
			return "\n".join(self.list[self.chelpg_charge_idxs[-1]:end_idx(self.chelpg_charge_idxs[-1] + 1)])
	@functools.lru_cache(maxsize=1)
	def _pop_analysis(self):
		if self.pop_analysis_idxs and self.muliken_charge_idxs:
			return "\n".join(self.list[self.pop_analysis_idxs[-1]:self.muliken_charge_idxs[-1]])
	@functools.lru_cache(maxsize=1)
	def _npa_analysis(self):
		if self.npa_start_idxs and self.npa_end_idxs:
			if len(self.npa_start_idxs) > 1:
				end_idx = lambda x: next(i for i, a in enumerate(self.list[x + 1:], start=x + 1) if a.strip() == "")
				return "\n".join(self.list[self.npa_start_idxs[-2]:end_idx(self.npa_end_idxs[-1])])
	@functools.lru_cache(maxsize=1)
	def _last_apt_charges(self):
		if self.apt_charge_idxs:
			end_idx = lambda x: next(i for i, a in enumerate(self.list[x + 1:], start=x + 1) if not a.split()[0].isdigit())
			return "\n".join(self.list[self.apt_charge_idxs[-1]:end_idx(self.apt_charge_idxs[-1]+1)])
	@functools.lru_cache(maxsize=1)
	def _raw_route(self):
		try:
			raw_route =None
			x = None if self.hash_line_idxs is None else min(a for a in self.multi_dash_idxs if a > self.hash_line_idxs[0])
			x = None if self.hash_line_idxs is None else "".join([a.lstrip() for a in self.list[self.hash_line_idxs[0]:x]])
			raw_route = " ".join(x.split())
		except IndexError as e:
			raw_route = None
			print("Error while finding route section of log file")
			print(e)
			print(self.name)
		finally:
			return raw_route
	@functools.lru_cache(maxsize=1)
	def _raw_route_keys(self):
		if not self.raw_route: return
		r_sect = [self.raw_route]
		for x in [None, "/", "(", ")", ",", "=", "%", ":"]:
			r_sect = [a for a in itertools.chain(*[i.split(x) for i in r_sect]) if len(a) > 1]
		r_sect = [a.lower() for a in r_sect]
		return r_sect

	def sep_conseq(self,mixed_list):
		x, group, new_list = None, [], []
		for a in mixed_list:
			if x is None and group == []: group.append(a); x = a
			elif x + 1 == a: group.append(a); x = a
			else: new_list.append(group); x = a; group = [a]
		new_list.append(group)
		return new_list

	@functools.lru_cache(maxsize=1)
	def _homo(self):
		try:
			error_a = f"Inconsisten orbitals on file\n{self.name}"
			assert all(min(a) == max(b) + 1 for a, b in zip(self.sep_conseq(self.uno_orb_energies), self.sep_conseq(self.oc_orb_energies))), error_a
			homo = []
			for structure in self.sep_conseq(self.oc_orb_energies):
				orbitals =[]
				for i in structure:
					error_b = f"Could not identify occupied orbital on line {i}\nFile:{self.name}"
					assert self.list[i].lstrip().startswith("Alpha  occ. eigenvalues --"), error_b
					for a in self.list[i].replace("Alpha  occ. eigenvalues --"," ").replace("-"," -").split():
						orbitals.append(float(a))
				homo.append(max(orbitals))
			return homo
		except AssertionError:
			return None
		except ValueError as e:
			print(f"Error while looking for homo energy on file {self.name}\n{e}")

	@functools.lru_cache(maxsize=1)
	def _lumo(self):
		try:
			error_a = f"Inconsisten orbitals on file\n{self.name}"
			assert all(min(a) == max(b) + 1 for a, b in zip(self.sep_conseq(self.uno_orb_energies), self.sep_conseq(self.oc_orb_energies))), error_a
			lumo = []
			for structure in self.sep_conseq(self.uno_orb_energies):
				orbitals = []
				for i in structure:
					error_b = f"Could not identify unoccupied orbital on line {i}\nFile:{self.name}"
					assert self.list[i].lstrip().startswith("Alpha virt. eigenvalues --"), error_b
					for a in self.list[i].replace("Alpha virt. eigenvalues --", " ").replace("-", " -").split():
						orbitals.append(float(a))
				lumo.append(min(orbitals))
			return lumo
		except AssertionError:
			return None
		except ValueError as e:
			print(f"Error while looking for lumo energy on file {self.name}\n{e}")
	@functools.lru_cache(maxsize=1)
	def	_homolumo(self):
		if self.homo and self.lumo:
			try:
				assert len(self.homo) == len(self.lumo), f"Inconsistent orbitals on file:\n{self.name}"
				return [a-b for a,b in zip(self.homo,self.lumo)]
			except AssertionError:
				return None

	homo = property(_homo)
	lumo = property(_lumo)
	homolumo = property(_homolumo)
	raw_route_keys = property(_raw_route_keys)
	raw_route = property(_raw_route)
	n_atoms = property(_n_atoms)
	normal_termin = property(_normal_termin)
	calc_type = property(_calc_type)
	error_msg = property(_error_msg)
	last_log_abstract = property(_last_log_abstract)
	last_freq = property(_last_freq)
	xyz_from_dist_matrix = property(_xyz_from_dist_matrix)
	last_muliken_spin_density = property(_last_muliken_spin_density)
	last_internal_coord = property(_last_internal_coord)
	last_muliken_charges = property(_last_muliken_charges)
	last_chelpg_charges = property(_last_chelpg_charges)
	pop_analysis = property(_pop_analysis)
	npa_analysis = property(_npa_analysis)
	last_apt_charges = property(_last_apt_charges)

class LogAbstract:
	def __init__(self,content):
		assert type(content) is list
		self.list = content
		self.version = None
		self.dipole = None
		self.img_freq = None
		self.hash_line = None
		for i,a in enumerate(self.list):
			a = a.lstrip()
			if a.lstrip == "": print("Empty!");continue
			elif a.startswith("Version="): self.version = a.replace("Version=","")
			elif a.startswith("#"): self.hash_line = i
			elif a.startswith("NImag="): self.img_freq = a.replace("NImag=0","")
			elif a.startswith("DipoleDeriv="): self.img_freq = a.replace("DipoleDeriv=","")
			else: continue
	def __str__(self):
		return "\n".join(self.list)
	def read_strucure(self):
		charge_mult = None
		title = None
		coordinates = []
		for i,a in enumerate(self.list[self.hash_line:]):
			if i > 5 and not coordinates: break
			a = a.split(",")
			if len(a) == 2 and not coordinates:	charge_mult = a; continue
			if len(a) == 4:
				if a[0] in elements and all(is_str_float(a[n]) for n in [1,2,3]):
					coordinates.append("   ".join(a))
				elif coordinates: break
			elif coordinates: break
		return charge_mult, XyzFile([self.list[0],str(len(coordinates)),title,*coordinates])
	def charge_mult(self):
		return self.read_strucure()[0]
	def xyz_object(self):
		return self.read_strucure()[1]
class LogFreq:
	def __init__(self, content):
		assert type(content) is list
		self.list = content
		self.rows = []
		for i,a in enumerate(self.list):
			if a.lstrip().startswith("Frequencies --"):
				try:
					assert self.list[i + 1].lstrip().startswith("Red. masses --")
					assert self.list[i + 2].lstrip().startswith("Frc consts  --")
					assert self.list[i + 3].lstrip().startswith("IR Inten    --")
					self.rows.append(i-2)
				except AssertionError:
					continue
		if not self.rows:
			self.n_atoms = 1
			self.block = []
		else:
			self.n_atoms = len(self.list) - self.rows[-1] - 7
			self.block = self.list[self.rows[0]:]
	def __str__(self):
		return "\n".join(self.list)
	@functools.lru_cache(maxsize=1)
	def frequencies(self):
		return list(itertools.chain(*[i.split()[2:] for i in self.block[2::self.n_atoms+7]]))
	@functools.lru_cache(maxsize=1)
	def ir_intensities(self):
		return list(itertools.chain(*[i.split()[3:] for i in self.block[5::self.n_atoms+7]]))
	@functools.lru_cache(maxsize=1)
	def displ_for_freq_idx(self,freq_idx):
		displ = []
		for num in range(self.n_atoms):
			displ.append(list(itertools.chain(*[i.split()[2:] for i in self.block[7+num::self.n_atoms+7]])))
		displ_for_freq_str = [a[freq_idx*3:freq_idx*3+3] for a in displ]
		displ_for_freq_float = [[float(i) for i in b] for b in displ_for_freq_str]
		return displ_for_freq_float
	@functools.lru_cache(maxsize=1)
	def n_ifreq(self):
		return str(len([b for b in self.frequencies() if float(b) < 0])) if self.frequencies() else "No data"
	def ir_spectra(self,threshold = 20):
		pairs = []
		for a,b in zip(self.frequencies(), self.ir_intensities()):
			if is_str_float(a) and is_str_float(b):
				pairs.append([float(a),float(b)])
		for a,b in zip(sorted(pairs,key=lambda x: x[1], reverse=True), range(threshold)):
			print("{:>10.1f}{:>10.1f}".format(float(a[0]),float(a[1])))
		print("---------")
class DistMatrix:
	def __init__(self,text):
		self.element = {}
		for a in [b.split() for b in text]:
			idx = "".join(a[0:2])
			if len(a) > 2 and a[1] in elements:
				if idx in self.element:
					self.element[idx].extend([float(c) for c in a[2:]])
					continue
				else:
					self.element[idx] = [float(c) for c in a[2:]]
		for a in self.element:
			print(a,self.element[a])
		self.dist_matrix = sorted(self.element.values(),key=lambda x: len(x))
		self.elem_vector = sorted(self.element.keys(), key=lambda x: len(self.element[x]))
		self.xyz_ent_a = []
		self.xyz_ent_b = []
		for i,a in enumerate(self.dist_matrix):
			if i == 0:
				self.xyz_ent_a.append([0, 0, 0])
				self.xyz_ent_b.append([0, 0, 0])
			if i == 1:
				self.xyz_ent_a.append([a[0], 0, 0])
				self.xyz_ent_b.append([a[0], 0, 0])
			if i == 2:
				x = (self.dist_matrix[i-1][0]**2+a[0]**2-a[1]**2)/(2*self.dist_matrix[i-1][0])
				y = math.sqrt(a[1]**2-x**2)
				self.xyz_ent_a.append([x, y, 0])
				self.xyz_ent_b.append([x, y, 0])
			if i > 2:
				x = (self.dist_matrix[i-1][0]**2+a[0]**2-a[1]**2)/(2*self.dist_matrix[i-1][0])
				y = math.sqrt(a[1]**2-x**2)
				#z =
				self.xyz_ent_a.append([x, y, 0])
				self.xyz_ent_b.append([x, y, 0])
class GjfFile:
	pattern = re.compile(r"[0-9][0-9][Gg]\+")
	def __init__(self,file_content):
		self.list = file_content
		self.list_l = [a.split() for a in file_content]
		self.str_l = [a.replace(" ", "") for a in self.list]
		self.return_print = "\n".join(self.list[1:])
		#########################
		#########################
		self.empty_line_idxs = [i for i,a in enumerate(self.list) if a.split() == []]
		self.asterisk_line_idxs = [idx for idx,line in enumerate(self.list) if line.split() == ["****"]]
		self.link_one_idxs = [i for i,l in enumerate(self.list) if "--link1--" in l.lower()]

	@functools.lru_cache(maxsize=1)
	def name(self):
		if len(self.list[0]) == 0: raise Exception(".gjf or .com object has no name")
		assert type(self.list[0]) is str, "Name must be string"
		return self.list[0]
	@functools.lru_cache(maxsize=1)
	def charge(self):
		return int(self.list[self.c_m_idx()].split()[0])
	@functools.lru_cache(maxsize=1)
	def multiplicity(self):
		return int(self.list[self.c_m_idx()].split()[1])
	@functools.lru_cache(maxsize=1)
	def n_electrons(self):
		return sum(elements.index(e) for e in self.all_elements()) - self.charge()
	@functools.lru_cache(maxsize=1)
	def n_atoms(self):
		return len(self.all_elements())
	@functools.lru_cache(maxsize=1)
	def all_elements(self):
		return [line[0] for line in self.cord_block()]
	@functools.lru_cache(maxsize=1)
	def elements(self):
		return list(dict.fromkeys(self.all_elements()))
	@functools.lru_cache(maxsize=1)
	def c_m_validate(self):
		return not self.n_electrons()%2 == self.multiplicity()%2
	@functools.lru_cache(maxsize=1)
	def c_m_validate_txt(self):
		return "Yes" if self.c_m_validate() else "--NO!--"
	@functools.lru_cache(maxsize=1)
	def n_proc(self):
		for line in self.list:
			line = line.lower().replace(" ","")
			if "%nprocshared=" in line:	return int(line.replace("%nprocshared=",""))
			elif "%nproc=" in line:	return int(line.replace("%nproc=",""))
	@functools.lru_cache(maxsize=1)
	def cord_block(self):
		coordinates = []
		for line in self.list_l[self.c_m_idx()+1:]:
			if len(line) == 0: break
			if len(line) != 4: continue
			if line[0] in elements:	coordinates.append(line)
			else: coordinates.append([elements[int(line[0])],*line[0:]])
		return coordinates
	@functools.lru_cache(maxsize=1)
	def route_text(self):
		flatten = lambda l: [item for sublist in l for item in sublist]
		try: return " ".join(flatten([a.split() for a in self.list[self.route_idx():self.title_idx()]]))
		except: return "No data"
	@functools.lru_cache(maxsize=1)
	def c_m_idx(self):
		if len(self.list[self.title_idx()+2].split()) < 2:
			raise Exception("Did you provide charge and multiplicity data at line {} of file {}?".format(self.title_idx()+1,self.name()))
		return self.title_idx()+2
	@functools.lru_cache(maxsize=1)
	def end_cord_idx(self):
		for idx,line in enumerate(self.list):
			if idx < self.c_m_idx(): continue
			if line.split() == []: return idx+1
	#########################
	#########################
	@functools.lru_cache(maxsize=1)
	def route_idx(self):
		for idx,line in enumerate(self.list):
			if line.strip().startswith("#"):return idx
		raise Exception("A route section (#) should be specified for .gjf or .com files")
	@functools.lru_cache(maxsize=1)
	def title_idx(self):
		for idx,line in enumerate(self.list):
			if idx > self.route_idx() and line.split() == []: return idx+1
	@functools.lru_cache(maxsize=1)
	def gen_basis(self):
		return any(i in self.route_text().lower() for i in ["/gen", "gen ","genecp"])
	@functools.lru_cache(maxsize=1)
	def declared_basis_lines(self):
		if not self.gen_basis(): return None
		idxs = [i+1 for idx,i in enumerate(self.asterisk_line_idxs) if i < self.asterisk_line_idxs[-1]]
		idxs.insert(0,max(i+1 for i in self.empty_line_idxs if  i < self.asterisk_line_idxs[-1]))
		return idxs
	@functools.lru_cache(maxsize=1)
	def declared_basis(self):
		e_w_b = [self.list[i].split()[:-1] for i in self.declared_basis_lines()]
		return [j.capitalize() for i in e_w_b for j in i]
	@functools.lru_cache(maxsize=1)
	def basis_errors(self):
		if not self.gen_basis(): return []
		#errors
		zero_last = any(self.list[i].split()[-1] == "0" for i in self.declared_basis_lines())
		miss_basis = [a for a in self.elements() if a not in self.declared_basis()]
		surpl_basis = [a for a in self.declared_basis() if a not in self.elements()]
		rep_basis = list(dict.fromkeys([a for a in self.declared_basis() if self.declared_basis().count(a) > 1]))
		errors = []
		for i in [*[a+1 for a in self.declared_basis_lines()],self.route_idx()]:
			if GjfFile.pattern.search(self.list[i]):
				errors.append("Is the basis set specifications correct?".format(i))
				errors.append("{}".format(self.list[i]))
				errors.append("Shouldn't '+' appear before the letter 'G'?")
		#statements
		if not zero_last:errors.append("Missing zero at the end of basis set specification?")
		if miss_basis:errors.append("Missing basis for: {} ?".format(" ".join(miss_basis)))
		if surpl_basis:errors.append("Surplous basis for: {} ?".format(" ".join(surpl_basis)))
		if rep_basis:errors.append("Repeated basis for: {} ?".format(" ".join(rep_basis)))
		return errors
	@functools.lru_cache(maxsize=1)
	def gen_ecp(self):
		return any(i in self.route_text().lower() for i in ["pseudo", "genecp"])
	@functools.lru_cache(maxsize=1)
	def declared_ecp_lines(self):
		line_idx = []
		if not self.gen_ecp(): return None
		if self.gen_basis(): start_idx = self.declared_basis_lines()[-1] + 1
		else:start_idx = self.end_cord_idx()
		for idx,line in enumerate(self.list):
			if idx < start_idx: continue
			if len(line.split()) <= 1: continue
			if line.split()[-1] != "0": continue
			if all(True if a.capitalize() in elements else False for a in line.split()[:-1]): line_idx.append(idx)
		return line_idx
	@functools.lru_cache(maxsize=1)
	def declared_ecp(self):
		ecps = [self.list[i].split()[:-1] for i in self.declared_ecp_lines()]
		return [j.capitalize() for i in ecps for j in i]
	@functools.lru_cache(maxsize=1)
	def ecp_errors(self,heavy_e = 36):
		if not self.gen_ecp(): return []
		#errors
		zero_last = any(self.list[i].split()[-1] == "0" for i in self.declared_ecp_lines())
		miss_ecp = [a for a in self.elements() if a not in self.declared_ecp() and elements.index(a) > heavy_e]
		surpl_ecp = [a for a in self.declared_ecp() if a not in self.elements()]
		rep_ecp = list(dict.fromkeys([a for a in self.declared_ecp() if self.declared_ecp().count(a) > 1]))
		#statements
		errors = []
		if not zero_last:errors.append("Missing zero at the end of ecp set specification?")
		if miss_ecp:errors.append("Missing ecp for: {} ?".format(" ".join(miss_ecp)))
		if surpl_ecp:errors.append("Surplous ecp for: {} ?".format(" ".join(surpl_ecp)))
		if rep_ecp:errors.append("Repeated ecp for: {} ?".format(" ".join(rep_ecp)))
		return errors
	@functools.lru_cache(maxsize=1)
	def route_errors(self):
		errors = []
		keywords = self.route_text().lower().split()
		if len(keywords) > 1:
			if "nosymm" in 	keywords:
				if keywords[0] == "#t" or keywords[0:2] == ["#","t"]:
					errors.append("Combination of 'NoSymm' and '#T' might supress geometry output!")
		return errors


	@functools.lru_cache(maxsize=1)
	def mem(self):
		for line in self.list:
			line = line.lower().replace(" ","")
			if line.startswith("%mem=") and line.endswith("mb"): return int(line[5:-2])
			elif line.startswith("%mem=") and line.endswith("gb"): return 1000*int(line[5:-2])
		return None
	#########################
	#########################
	def replace_cord(self, xyz_obj):
		new = []
		for line in self.list[0:self.c_m_idx() + 1]: new.append(line)
		for line in xyz_obj.form_cord_block(): new.append(line)
		for line in self.list[self.end_cord_idx()-1:]: new.append(line)
		return GjfFile(new)
	def xyz_obj(self):
		return XyzFile([self.name(),self.n_atoms()," ",*[" ".join(a) for a in self.cord_block()]])
class XyzFile:
	def __init__(self,file_content):
		self.list = file_content
		if len(self.list) < 2: raise Exception(".xyz Object is empty?")
		elif not (str(self.list[1]).strip().isdigit() and len(str(self.list[1]).split()) == 1):
			print("{} is not a proper .xyz file\nAttempting to read it anyway!".format(self.list[0]))
			try_xyz = []
			for line in self.list:
				line = line.split()
				if len(line) != 4: continue
				if not all(is_str_float(line[i]) for i in range(1, 4)): continue
				if line[0] in elements[0:]:
					try_xyz.append(" ".join(line))
					continue
				try:
					line[0] = elements[int(line[0])]
					try_xyz.append(" ".join(line))
				except:
					raise Exception("Could not understand file {}".format(self.list[0]))
			try_xyz.insert(0,len(try_xyz))
			try_xyz.insert(1," ")
			try_xyz.insert(0,self.list[0])
			self.list = try_xyz
		self.list_l = [str(a).split() for a in self.list]
		#self.molecule.print_int_bond_map()
	def __add__(self,other):
		assert type(self) == type(other), "Operation '+' allowed only for two XYZ objects"
		new = [os.path.splitext(self.name())[0]+"_"+other.name(), str(self.n_atoms()+other.n_atoms()),
			   self.title()+" "+other.title(),*(self.form_cord_block() + other.form_cord_block())]
		return XyzFile(new)
	def __sub__(self, other):
		el_a = self.all_elements()
		el_b = other.all_elements()
		assert len(el_a) > len (el_b), "Can't subtract a larger structure from a smaller one"
		assert type(self) == type(other), "Operation '-' allowed only for two XYZ objects"
		idxs_to_rem = []
		for n in range(len(el_a) - len(el_b)):
			if all([True if el_a[n+i] == a else False for i,a in enumerate(el_b)]):
				idxs_to_rem = [c+n for c in range(len(el_b))]
				break
		if len(idxs_to_rem) ==  0: print("Could not subtract value!")
		xyz_cord = [a for idx,a in enumerate(self.form_cord_block()) if idx not in idxs_to_rem]
		new = [os.path.splitext(self.name())[0]+"-"+other.name(), str(self.n_atoms()-other.n_atoms()),
			   self.title()+"-"+other.title(),*xyz_cord]
		return XyzFile(new)
	def __str__(self):
		return "\n".join(self.return_print())
	@functools.lru_cache(maxsize=1)
	def name(self):
		if len(self.list[0]) == 0: raise Exception(".xyz Object has no name")
		return self.list[0]
	@functools.lru_cache(maxsize=1)
	def n_atoms(self):
		if any([len(str(self.list[1]).split()) != 1, not str(self.list[1]).isnumeric()]):
			raise Exception("First line of {} (.xyz type) file should contain only the number of atoms in the geometry!".format(self.name()))
		return int(self.list[1])
	@functools.lru_cache(maxsize=1)
	def title(self):
		return self.list[2]
	@functools.lru_cache(maxsize=1)
	def cord_block(self):
		cordinates = []
		for idx,line in enumerate(self.list_l):
			if idx <= 2: continue
			if idx >= self.n_atoms() + 3: continue
			if line[0] in elements:	cordinates.append(line)
			else: cordinates.append([elements[int(line[0])],*line[0:]])
		return cordinates
	@functools.lru_cache(maxsize=1)
	def form_cord_block(self):
		return ["{:<5}{:>20.6f}{:>20.6f}{:>20.6f}".format(x[0], *[float(x[a]) for a in [1, 2, 3]]) for x in self.cord_block()]
	@functools.lru_cache(maxsize=1)
	def cord_strip(self):
		return [line[1:] for line in self.cord_block()]
	@functools.lru_cache(maxsize=1)
	def all_elements(self):
		return [line[0] for line in self.cord_block()]
	@functools.lru_cache(maxsize=1)
	def elements(self):
		return list(dict.fromkeys(self.all_elements()))
	@functools.lru_cache(maxsize=1)
	def n_electrons(self):
		return sum(elements.index(e) for e in self.all_elements())
	@functools.lru_cache(maxsize=1)
	def return_print(self):
		return [str(self.n_atoms()),self.title(),*[l for l in self.form_cord_block()]]
	def print_file(self):
		print("======={}=======".format(self.name()))
		print("=======START=======")
		print("\n".join([l for l in self.return_print()]))
		print("========END========")
	def save_file(self,directory=None):
		if directory is None:
			file_path = os.path.splitext(os.path.join(os.getcwd(),self.name().replace(" ","")))[0]+".xyz"
		else:
			file_path = os.path.splitext(os.path.join(directory,self.name().replace(" ","")))[0]+".xyz"
		if os.path.exists(file_path):
			print("File {} already exists!".format(os.path.splitext(os.path.basename(file_path))[0] + ".xyz"))
			return
		with open(file_path,"w") as file:
			for line in self.return_print():file.write(str(line)+"\n")
		print("File {} saved!".format(os.path.splitext(os.path.basename(file_path))[0] + ".xyz"))
	def print_all(self):
		print("\n".join([l for l in self.list]))
	def displace(self,mult,displacement):
		cord_block = [[a,*[float(b[n])-c[n]*mult for n in range(3)]] for a,b,c in zip(self.all_elements(),self.cord_strip(),displacement)]
		cord_block = [" ".join([str(i) for i in l]) for l in cord_block]
		return XyzFile([self.name(),self.n_atoms(),self.title(),*cord_block])
	def rotate(self, angle, axis):
		"takes xyz object and returns xyz object rotated by angle over axis"
		assert axis in ("x", "y", "z"), "Only 'x','y' or 'z' axis are suported"
		if axis == "x":
			m_mat = [[1., 0., 0.], [0., math.cos(angle), -math.sin(angle)], [0., math.sin(angle), math.cos(angle)]]
		if axis == "y":
			m_mat = [[math.cos(angle), 0., math.sin(angle)], [0., 1., 0.], [-math.sin(angle), 0., math.cos(angle)]]
		if axis == "z":
			m_mat = [[math.cos(angle), -math.sin(angle), 0.], [math.sin(angle), math.cos(angle), 0.], [0., 0., 1.]]
		m_mat = np.array(m_mat, np.float64)
		rotated = np.array([i[1:4] for i in self.cord_block()], np.float64).transpose()
		rotated = np.matmul(m_mat,rotated).transpose()
		rotated = np.ndarray.tolist(rotated)
		rotated = [[b,*[str(n) for n in a]] for b,a in zip(self.all_elements(),rotated)]
		xyz_mat = [self.name(), self.n_atoms()," ",*[" ".join(a) for a in rotated]]
		return XyzFile(xyz_mat)
	def superimpose(self, other, num_atoms=0, print_step=False, ret = "geom",conv=18):
		"""Takes xyz object and returns xyz object rotated by angle over axis.
		Returns either the max_distance 'max_d' or final geometry 'geom' after rotations and superpositions"""
		def rotate(xyz,angle,axis):
			assert axis in ("x","y","z"), "Only 'x','y' or 'z' axis are suported"
			if axis == "x":
				m_mat = [[1., 0., 0.], [0., math.cos(angle), -math.sin(angle)], [0., math.sin(angle), math.cos(angle)]]
			if axis == "y":
				m_mat = [[math.cos(angle), 0., math.sin(angle)], [0., 1., 0.], [-math.sin(angle), 0., math.cos(angle)]]
			if axis == "z":
				m_mat = [[math.cos(angle), -math.sin(angle), 0.], [math.sin(angle), math.cos(angle), 0.], [0., 0., 1.]]
			m_mat = np.array(m_mat, np.float64)
			rotated = np.array(xyz, np.float64).transpose()
			rotated = np.matmul(m_mat,rotated).transpose()
			return np.ndarray.tolist(rotated)
		def calc_err(xyz_1, xyz_2, n_atms):
			n_atms = len(xyz_1) if n_atms == 0 else n_atms
			sq_dist = sum(sum(math.pow(c-d,2) for c,d in zip(a,b)) for a,b in zip(xyz_1[:n_atms],xyz_2))
			return math.sqrt(sq_dist / n_atms)
		def max_dist(xyz_a, xyz_b):
			return max(math.sqrt(sum(pow(c-d,2) for c,d in zip(a,b))) for a,b in zip(xyz_a,xyz_b))
		#----------------------
		last_error = None
		xyz_1 = [[float(a) for a in b] for b in other.std_cord(num_atoms).cord_strip()]
		xyz_2 = [[float(a) for a in b] for b in self.std_cord(num_atoms).cord_strip()]
		#Check atom correspondence
		for a,b,c in zip(range(len(self.all_elements()) if num_atoms == 0 else num_atoms),other.all_elements(),self.all_elements()):
			if b != c:
				atom_number = 'th' if 11<=a+1<=13 else {1:'st',2:'nd',3:'rd'}.get((a+1)%10, 'th')
				print("WARNING: {}{} atom pair doesn't not correspond to an element match: {} & {}".format(a+1,atom_number,b,c))
		if print_step: print("======ACTIONS======")
		#Start algorithm
		for num in range(conv):
			step_size = 1 / 2 ** num
			while True:
				rot = [[1, "x"], [1, "y"], [1, "z"], [-1, "x"], [-1, "y"], [-1, "z"]]
				movements = [rotate(xyz_2, step_size * i[0], i[1]) for i in rot]
				if ret == "max_d":
					last_error = max_dist(xyz_2, xyz_1)
					errors = [max_dist(i, xyz_1) for i in movements]
				else:
					last_error = calc_err(xyz_2, xyz_1, num_atoms)
					errors = [calc_err(i, xyz_1, num_atoms) for i in movements]
				best_m = errors.index(min(errors))
				if min(errors) < last_error:
					xyz_2 = movements[best_m]
					if print_step:
						msg = [step_size * rot[best_m][0], rot[best_m][1], calc_err(xyz_1, xyz_2, num_atoms)]
						print("Rotating {:.5f} radian in {}. RMSD = {:.5f}".format(*msg))
					continue
				else:
					if ret == "max_d" and max_dist(xyz_1, xyz_2) < 0.1:
						return True
					break
		if print_step: print("Final RMSD = {:.5f}".format(calc_err(xyz_1, xyz_2, num_atoms)))
		if print_step: print("========END========")
		if ret == "geom":
			cord_block = [" ".join([a,*[str(n) for n in b]]) for a,b in zip(self.all_elements(),xyz_2)]
			return XyzFile([self.name(),self.n_atoms(),self.title(),*cord_block])
		elif ret == "max_d":
			return False
	def std_cord(self, n_atoms=0):
		pure_cord = self.cord_strip() if n_atoms == 0 else self.cord_strip()[0:n_atoms]
		xyz_avg = [[float(n) for n in i] for i in pure_cord]
		xyz_avg = [sum([i[n] for i in xyz_avg]) / len(xyz_avg) for n in range(3)]
		xyz_avg = [[float(i[n]) - xyz_avg[n] for n in range(3)] for i in self.cord_strip()]
		xyz_avg = [[str(n) for n in a] for a in xyz_avg]
		xyz_avg = [" ".join([b,*a]) for b,a in zip(self.all_elements(),xyz_avg)]
		xyz_mat = [self.name(), self.n_atoms(), " ", *xyz_avg]
		return XyzFile(xyz_mat)
	def enantiomer(self):
		xyz = [" ".join([*a[0:-1],str(-float(a[-1]))]) for a in self.cord_block()]
		xyz_mat = [os.path.splitext(self.name())[0]+"_ent.xyz", self.n_atoms(), " ", *xyz]
		return XyzFile(xyz_mat)
	molecule = property(lambda self: Molecule(self.cord_block()))
class Molecule:
	def __init__(self,atom_list):
		assert type(atom_list) is list
		self.atom_list = [Atom(a,i) for i,a in enumerate(atom_list)]
		self.abc_angle(0,1,2)
		self.n_mol_ent()
	def __str__(self):
		return "\n".join([str(a) for a in self.atom_list])
	def int_bond_map(self):
		return [[b.int_bond_order(a) if a != b and b.int_bond_order(a) > 0.85 else None for a in self.atom_list] for b in self.atom_list]
	def ts_bond_map(self):
		return [[b.ts_bond_order(a) if a != b and b.int_bond_order(a) > 0.85 else None for a in self.atom_list] for b in self.atom_list]
	def print_int_bond_map(self):
		for a in self.atom_list:
			bonded = ["{:>3}{:>2}:{:.1f}".format(b.idx,b.element, b.int_bond_order(a)) for b in self.atom_list if a != b and b.int_bond_order(a) > 0.1]
			print("{:>3}{:>2}".format(a.idx,a.element),"-->",", ".join(bonded))
	def print_ts_bond_map(self):
		for a in self.atom_list:
			bonded = ["{:>3}{:>2}:{:.1f}".format(b.idx,b.element, b.ts_bond_order(a)) for b in self.atom_list if a != b and b.ts_bond_order(a) > 0.1]
			print("{:>3}{:>2}".format(a.idx,a.element),"-->",", ".join(bonded))
	def n_mol_ent(self, map=None):
		if map is None: map = self.int_bond_map()
		visited = [False for _ in map]
		n_entities = 0
		entities = []
		def check(idx,atoms=[]):
			visited[idx] = True
			atoms.append(idx)
			for ib, b in enumerate(map[idx]):
				if b is None: continue
				elif visited[ib]: continue
				else:
					print(f"Leaving {idx+1} to check on {ib+1} because of BO: {b}")
					check(ib, atoms)
			return atoms
		for ia,a in enumerate(map):
			if visited[ia]: continue
			else:
				print(f"Adding new entitie starting from {ia+1}")
				n_entities +=1
				entities.append(check(ia,[]))
		print("Visited\n",visited)
		print("n entitites\n", n_entities)
		print("entities\n", entities)

	def valid_idxs(func):
		def wrapper(obj,*list):
			assert all([type(n) is int for n in list]), "Atom indexes should be integers"
			assert all([n in range(len(obj.atom_list)) for n in list]), "Atom indexes are out of range"
			return func(obj,*list)
		return wrapper
	@valid_idxs
	def ab_distance(self,a,b):
		return self.atom_list[a].distance(self.atom_list[b])
	@valid_idxs
	def abc_angle(self,a,b,c):
		return self.atom_list[a].angle(self.atom_list[b],self.atom_list[c])
	@valid_idxs
	def abcd_dihedral(self,a,b,c,d):
		return self.atom_list[a].dihedral(self.atom_list[b],self.atom_list[c],self.atom_list[d])
class Atom:
	el_radii = dict(element_radii)
	def __init__(self,line,idx):
		assert type(line) is list
		assert len(line) == 4
		assert line[0] in elements
		assert all(is_str_float(a) for a in line[1:])
		self.idx = idx
		self.element = line[0]
		self.cord = [float(a) for a in line[1:]]
	def distance(self,other):
		return sum((b - a) ** 2 for a, b in zip(self.cord, other.cord)) ** 0.5
	def angle(self,other_a,other_b):
		a_a = np.array(self.cord)
		b_a = np.array(other_a.cord)
		c_a = np.array(other_b.cord)
		ba, bc = a_a - b_a, c_a - b_a
		cosine_angle = np.dot(ba, bc) / (np.linalg.norm(ba) * np.linalg.norm(bc))
		angle = np.arccos(cosine_angle)
		#print("Angle :",self.idx,other_a.idx,other_b.idx,"is:", "{:.2f}°".format(np.degrees(angle)))
		return angle
	def dihedral(self,other_a,other_b,other_c):
		p = np.array([self.cord,other_a.cord,other_b.cord,other_c.cord])
		# From: stackoverflow.com/questions/20305272/dihedral-torsion-angle-from-four-points-in-cartesian-coordinates-in-python
		b = p[:-1] - p[1:]
		b[0] *= -1
		v = np.array([np.cross(v, b[1]) for v in [b[0], b[2]]])
		# Normalize vectors
		v /= np.sqrt(np.einsum('...i,...i', v, v)).reshape(-1, 1)
		return np.degrees(np.arccos(v[0].dot(v[1])))
	def ts_bond_order(self,other):
		return math.exp((Atom.el_radii[self.element]/100 + Atom.el_radii[other.element]/100 - self.distance(other))/0.6)
	def int_bond_order(self,other):
		return math.exp((Atom.el_radii[self.element]/100 + Atom.el_radii[other.element]/100 - self.distance(other))/0.3)
	def __str__(self):
		return "{}{}".format(self.idx,self.element)

class Var:
	conf_dir = os.path.dirname(__file__)
	conf_file = os.path.join(conf_dir, "chemxls_preferences.init")
	def __init__(self,conf_file=conf_file):
		self.ext = ["any", ".xyz", ".gjf", ".com", ".log", ".inp", ".out"]
		a = self.ext
		self.options = [
			{"short":"Blank",              "uid":"001", "extension":a[0], "supl":False, "float":False, "hyp":False, "long":"Blank column"                                               },
			{"short":"Eh to kcal/mol",     "uid":"002", "extension":a[0], "supl":False, "float":True , "hyp":False, "long":"Hartree to kcal/mol conversion factor (627.5)"              },
			{"short":"Eh to kJ/mol",       "uid":"003", "extension":a[0], "supl":False, "float":True , "hyp":False, "long":"Hartree to kJ/mol conversion factor (2625.5)"               },
			{"short":"Filename",           "uid":"004", "extension":a[0], "supl":False, "float":False, "hyp":False, "long":"Filename"                                                   },
			{"short":"Folder",             "uid":"005", "extension":a[0], "supl":False, "float":False, "hyp":True , "long":"Hyperlink to corresponding folder"                          },
			{"short":".xyz",               "uid":"006", "extension":a[1], "supl":False, "float":False, "hyp":True , "long":"Hyperlink to Filename.xyz"                                  },
			{"short":".gjf",               "uid":"007", "extension":a[2], "supl":False, "float":False, "hyp":True , "long":"Hyperlink to Filename.gjf"                                  },
			{"short":".gjf_#",             "uid":"008", "extension":a[2], "supl":False, "float":False, "hyp":False, "long":"Route section read from Filename.gjf"                       },
			{"short":".com",               "uid":"009", "extension":a[3], "supl":False, "float":False, "hyp":True , "long":"Hyperlink to Filename.com"                                  },
			{"short":".com_#",             "uid":"010", "extension":a[3], "supl":False, "float":False, "hyp":False, "long":"Route section read from Filename.com"                       },
			{"short":".log",               "uid":"011", "extension":a[4], "supl":False, "float":False, "hyp":True , "long":"Hyperlink to Filename.log"                                  },
			{"short":".log_#",             "uid":"012", "extension":a[4], "supl":False, "float":False, "hyp":False, "long":"Route section read from Filename.log"                       },
			{"short":"E0",                 "uid":"013", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Energy from last SCF cycle"                                 },
			{"short":"iFreq",              "uid":"014", "extension":a[4], "supl":False, "float":False, "hyp":False, "long":"Number of imaginary frequencies found on Filename.log"      },
			{"short":"E_ZPE",              "uid":"015", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Zero-point correction"                                      },
			{"short":"E_tot",              "uid":"016", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Thermal correction to Energy"                               },
			{"short":"H_corr",             "uid":"017", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Thermal correction to Enthalpy"                             },
			{"short":"G_corr",             "uid":"018", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Thermal correction to Gibbs Free Energy"                    },
			{"short":"E0+E_ZPE",           "uid":"019", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Sum of electronic and zero-point Energies"                  },
			{"short":"E0+E_tot",           "uid":"020", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Sum of electronic and thermal Energies"                     },
			{"short":"E0+H_corr",          "uid":"021", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Sum of electronic and thermal Enthalpies"                   },
			{"short":"E0+G_corr",          "uid":"022", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Sum of electronic and thermal Free Energies"                },
			{"short":"Done?",              "uid":"023", "extension":a[4], "supl":False, "float":False, "hyp":False, "long":"Filename.log gaussian normal termination status"            },
			{"short":"Error",              "uid":"024", "extension":a[4], "supl":False, "float":False, "hyp":False, "long":"Error messages found on Filename.log"                       },
			{"short":"HOMO",               "uid":"025", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"HOMO from Alpha  occ. eigenvalues of Filename.log"          },
			{"short":"LUMO",               "uid":"026", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"LUMO from Alpha virt. eigenvalues of Filename.log"          },
			{"short":"HOMO-LUMO",          "uid":"027", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"HOMO-LUMO from Alpha occ. & virt. eigenv. of Filename.log"  },
			{"short":"Charge",             "uid":"028", "extension":a[4], "supl":False, "float":False, "hyp":False, "long":"Charge from Filename.log"                                   },
			{"short":"Mult",               "uid":"029", "extension":a[4], "supl":False, "float":False, "hyp":False, "long":"Starting multiplicity from Filename.log"                    },
			{"short":"n_SCF",              "uid":"030", "extension":a[4], "supl":False, "float":False, "hyp":False, "long":"Number of 'SCF Done:' keywords found"                       },
			{"short":"n_atoms",            "uid":"031", "extension":a[4], "supl":False, "float":False, "hyp":False, "long":"Number of atoms on Filename.log"                            },
			{"short":"TYP",                "uid":"032", "extension":a[4], "supl":False, "float":False, "hyp":False, "long":"Filename.log calculation type (This may be unreliable)"     },
			{"short":"Needs refinement?",  "uid":"033", "extension":a[4], "supl":False, "float":False, "hyp":False, "long":"Filename.log calculation type consistency with iFreq"       },
			{"short":"S**2 BA",            "uid":"034", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Filename.log last spin densities before anihilation"        },
			{"short":"S**2 After",         "uid":"035", "extension":a[4], "supl":False, "float":True , "hyp":False, "long":"Filename.log last spin densities after anihilation"         },
			{"short":"LG",                 "uid":"036", "extension":a[4], "supl":True , "float":False, "hyp":True , "long":"Filename.log last geometry"                                 },
			{"short":"MulkSpinDens",       "uid":"037", "extension":a[4], "supl":True , "float":False, "hyp":True , "long":"Filename.log last Muliken charge and spin density"          },
			{"short":"LastIntCoord",       "uid":"038", "extension":a[4], "supl":True , "float":False, "hyp":True , "long":"Filename.log last Internal coordinates"                     },
			{"short":"MulkCharges",        "uid":"039", "extension":a[4], "supl":True , "float":False, "hyp":True , "long":"Filename.log last Muliken charge"                           },
			{"short":"ESPCharges",         "uid":"040", "extension":a[4], "supl":True , "float":False, "hyp":True , "long":"Filename.log last ESP charge"                               },
			{"short":"POPAnalysis",        "uid":"041", "extension":a[4], "supl":True , "float":False, "hyp":True , "long":"Filename.log last population analysis"                      },
			{"short":"NPAAnalysis",        "uid":"042", "extension":a[4], "supl":True , "float":False, "hyp":True , "long":"Filename.log last NPA analysis"                             },
			{"short":"APTCharges",         "uid":"043", "extension":a[4], "supl":True , "float":False, "hyp":True , "long":"Filename.log last APT charges"                              },
			{"short":".inp",               "uid":"044", "extension":a[5], "supl":False, "float":False, "hyp":True ,	"long":"Hyperlink to Filename.inp"                                  },
			{"short":".out",               "uid":"045", "extension":a[6], "supl":False, "float":False, "hyp":True , "long":"Hyperlink to Filename.out"                                  },
		]

		assert not any(a["hyp"] and a["float"] for a in self.options), "Cannot be float and hyperlink simultaneously"
		assert len(set(a["uid"] for a in self.options)) == len(self.options), "UIDs have to be unique"
		assert len(set(a["short"] for a in self.options)) == len(self.options), "Short names have to be unique"
		assert len(set(a["long"] for a in self.options)) == len(self.options), "Long names have to be unique"
		assert set(a["extension"] for a in self.options) == set(self.ext), "Are there unused extensions or typos?"
		assert all([a["hyp"] and a["supl"] for a in self.options if a["supl"]]), "Use of suplementary files must be accompanied by corresponding hyperlink"

		self.std_config = configparser.ConfigParser()
		self.std_config["DEFAULT"] = {"options": "005 004 006 007 011","splitext":"False","splitjobs":"False"}
		self.std_config["STARTUP"] = {"options": "005 004 006 007 011","splitext":"False","splitjobs":"False"}
		self.std_config["PRESETA"] = {"options": "005 004 006 007"    ,"splitext":"False","splitjobs":"False"}
		self.std_config["PRESETB"] = {"options": "005 004 006"        ,"splitext":"False","splitjobs":"False"}
		self.std_config["PRESETC"] = {"options": "005 004"            ,"splitext":"False","splitjobs":"False"}
		if not os.path.isfile(conf_file):
			with open(conf_file, "w") as configfile:
				self.std_config.write(configfile)
			self.config = self.std_config
		else:
			self.config = configparser.ConfigParser()
			self.config.read(conf_file)
		def pick(args,get_type,valid_keys={},default=None,config=self.config,std_config=self.std_config):
			try:
				if   get_type == "str" : result = config.get(*args)
				elif get_type == "bool": result = config.getboolean(*args)
				elif get_type == "int" : result = config.getint(*args)
			except:
				if   get_type == "str" : result = std_config.get(*args)
				elif get_type == "bool": result = std_config.getboolean(*args)
				elif get_type == "int" : result = std_config.getint(*args)
			finally:
				if valid_keys          : result = valid_keys.get(result,default)
				return result
		big_name = ["DEFAULT"           ,"STARTUP"           ,"PRESETA"            ,"PRESETB"            ,"PRESETC"            ]
		opt      = ["default_options"   ,"startup_options"   ,"preset_a_options"   ,"preset_b_options"   ,"preset_c_options"   ]
		split    = ["default_split"     ,"startup_split"     ,"preset_a_split"     ,"preset_b_split"     ,"preset_c_split"     ]
		jobs     = ["default_split_jobs","startup_split_jobs","preset_a_split_jobs","preset_b_split_jobs","preset_c_split_jobs"]
		valid_keys = [a["uid"] for a in self.options]
		for big,a,b,c in zip(big_name,opt,split,jobs):
			setattr(self,a,[n for n in pick((big,"options"),"str").split() if n in valid_keys])
			setattr(self,b,pick((big,"splitext" ),"bool", default=False))
			setattr(self,c,pick((big,"splitjobs"),"bool", default=False))

	def set_variables(self,section,option,value,conf_file=conf_file):
		self.config[section][option] = value
		with open(conf_file, "w") as configfile:
			self.config.write(configfile)
		self.__init__()

#GUI CLASSES
class FileFolderSelection(tk.Frame):
	def __init__(self,parent):
		tk.Frame.__init__(self,parent)
		self.in_folder = None # str
		self.in_folder_label = tk.StringVar(value="Please set folder name")
		self.supl_folder = None # str
		self.supl_folder_label = tk.StringVar(value="Please set folder name")
		self.supl_folder_auto = tk.BooleanVar(value=True)
		self.xls_path = None # str
		self.xls_path_label = tk.StringVar(value="Please set file name")
		self.xls_path_auto = tk.BooleanVar(value=True)
		self.recursive_analysis = tk.BooleanVar(value=True)
		self.str_width = 500
		self.grid_columnconfigure(0, weight=1)
		self.lock = False
		self.check_buttons = []
		self.buttons = []

		#INPUT FOLDER
		box = self.boxify("Analyze this directory:", 0)
		label = tk.Label(box, textvariable=self.in_folder_label)
		label.config(width=self.str_width, fg="navy")
		label.grid(column=0, row=0)
		button = tk.Button(box, text="Select", command=self.set_in_folder, padx="1", pady="0")
		button.config(fg="navy")
		button.grid(column=1, row=0, sticky="e")
		self.buttons.append(button)
		check_button = tk.Checkbutton(box, text="Recursively",
									  variable=self.recursive_analysis,
									  onvalue=True,
									  offvalue=False,
									  selectcolor="gold")
		check_button.grid(column=2, row=0, sticky="w")
		self.check_buttons.append(check_button)

		#SUPLEMENTARY FOLDER
		box = self.boxify("Write suplementary files to this directory:", 1)
		label = tk.Label(box, textvariable=self.supl_folder_label)
		label.config(width=self.str_width)
		label.grid(column=0, row=0)
		button = tk.Button(box, text="Select", command=self.set_supl_folder, padx="1", pady="0")
		button.grid(column=1, row=0, sticky="e")
		self.buttons.append(button)
		check_button = tk.Checkbutton(box, text="Auto",
									  variable=self.supl_folder_auto,
									  onvalue=True, offvalue=False,
									  command=self.auto_set_supl)
		check_button.grid(column=2, row=0, sticky="w")
		self.check_buttons.append(check_button)
		# XLS file
		box = self.boxify("Write xls file here:", 2)
		label = tk.Label(box, textvariable=self.xls_path_label)
		label.config(width=self.str_width)
		label.grid(column=0, row=0)
		button = tk.Button(box, text="Select", command=self.set_xls_path, padx="1", pady="0")
		button.grid(column=1, row=0, sticky="e")
		self.buttons.append(button)
		check_button = tk.Checkbutton(box, text="Auto",
									  variable=self.xls_path_auto,
									  onvalue=True, offvalue=False,
									  command=self.auto_set_xls)
		check_button.grid(column=2, row=0, sticky="w")
		self.check_buttons.append(check_button)

		#AUTO SET
		if len(sys.argv) > 1 and sys.argv[-1] in ["--cwd","-cwd","cwd"]:
			self.in_folder = os.path.normpath(os.getcwd())
			self.in_folder_label.set(trim_str(self.in_folder,self.str_width))
			if self.xls_path_auto.get(): self.auto_set_xls()
			if self.supl_folder_auto.get(): self.auto_set_supl()

	def boxify(self,name,row):
		box = tk.LabelFrame(self, text=name)
		box.grid(column=0, row=row, sticky="news")
		box.grid_columnconfigure(2, minsize=90)
		box.grid_columnconfigure(0, weight=1)
		return box
	def set_in_folder(self):
		in_folder = filedialog.askdirectory()
		assert type(in_folder) == str
		if type(in_folder) == str and in_folder.strip() != "":
			self.in_folder = os.path.normpath(in_folder)
			self.in_folder_label.set(trim_str(self.in_folder,self.str_width))
			if self.xls_path_auto.get(): self.auto_set_xls()
			if self.supl_folder_auto.get(): self.auto_set_supl()
		else:
			messagebox.showinfo(title="Folder selection", message="Folder won't be set!")
	def set_supl_folder(self):
		supl_folder = filedialog.askdirectory()
		if type(supl_folder) == str and supl_folder.strip() != "":
			self.supl_folder = os.path.normpath(os.path.join(supl_folder, "chemxlslx_supl_files"))
			self.supl_folder_auto.set(False)
			self.supl_folder_label.set(trim_str(self.supl_folder, self.str_width))
		else:
			messagebox.showinfo(title="Folder selection", message="Folder won't be set!")

	def auto_set_supl(self):
		if self.supl_folder_auto.get() and type(self.in_folder) == str:
			if not os.path.isdir(self.in_folder): return
			supl_folder = os.path.join(self.in_folder, "chemxlslx_supl_files")
			supl_folder = os.path.normpath(supl_folder)
			self.supl_folder = supl_folder
			self.supl_folder_label.set(trim_str(supl_folder,self.str_width))
	def set_xls_path(self):
		xls_path = filedialog.asksaveasfilename(title = "Save xls file as:",
													  filetypes = [("Spreadsheet","*.xls")])
		assert type(xls_path) == str
		if os.path.isdir(os.path.dirname(xls_path)) and xls_path.strip() != "":
			if not xls_path.endswith(".xls"):	xls_path += ".xls"
			self.xls_path = os.path.normpath(xls_path)
			self.xls_path_auto.set(False)
			self.xls_path_label.set(trim_str(self.xls_path,self.str_width))
		else:
			messagebox.showinfo(title="File selection", message="File won't be set!")

	def auto_set_xls(self):
		if self.xls_path_auto.get() and type(self.in_folder) == str:
			if not os.path.isdir(self.in_folder): return
			self.xls_path = os.path.join(self.in_folder,"chemxls_analysis.xls")
			self.xls_path = os.path.normpath(self.xls_path)
			self.xls_path_label.set(trim_str(self.xls_path,self.str_width))

class ListBoxFrame(tk.Frame):
	def __init__(self,parent):
		tk.Frame.__init__(self,parent)
		self.root = parent
		self.grid_columnconfigure(0,weight=1)
		self.grid_columnconfigure(3,weight=1)
		self.columnconfigure(0,uniform="fred")
		self.columnconfigure(3,uniform="fred")
		self.grid_rowconfigure(1,weight=1)
		self.grid_rowconfigure(2,weight=1)
		self.lock = False
		self.preferences = Var()
		self.options = self.preferences.options
		self.need_style0 = [a["short"] for a in self.options if a["float"]]
		self.need_formula = [a["short"] for a in self.options if a["hyp"]]
		self.dict_options = {a["long"]:[a["short"],a["uid"],a["extension"]] for a in self.options}
		self.label_dict = {a["short"]:a["long"] for a in self.options}
		self.label_dict.update({"Link1":"Job step of 'Filename.log'"})
		self.extension_dict =  {d["short"]:d["extension"] for d in self.options}
		self.extension_dict.update({"Link1":".log"})
		#LEFT PANEL
		left_label = tk.Label(self,text="Available options")
		left_label.grid(column=0,row=0,columnspan=2)
		self.listbox_a = tk.Listbox(self)
		self.populate_a("any")
		self.listbox_a.grid(column=0, row=1,rowspan=4,sticky="news")
		scrollbar = tk.Scrollbar(self, orient="vertical")
		scrollbar.config(command=self.listbox_a.yview)
		scrollbar.grid(column=1, row=1, rowspan=4, sticky="ns")
		self.listbox_a.config(yscrollcommand=scrollbar.set)
		self.check_buttons = []
		self.buttons = []

		#BOTTOM LEFT PANEL
		frame = tk.Frame(self)
		frame.grid(column=0,row=5,columnspan=2)
		left_label = tk.Label(frame,text="Filter options by file extension:")
		left_label.grid(column=0,row=0)
		self.display_ext = tk.StringVar()
		self.drop_options = ttk.OptionMenu(frame,self.display_ext,"any",*self.preferences.ext,
										   command=lambda x=self.display_ext.get():self.populate_a(x))
		self.drop_options.configure(width=10)
		self.drop_options.grid(column=1,row=0,sticky="e")

		#RIGHT PANEL
		right_label = tk.Label(self,text="Selected options")
		right_label.grid(column=3,row=0,columnspan=2)
		self.listbox_b = tk.Listbox(self)
		self.listbox_b.grid(column=3, row=1, rowspan=4, sticky="news")
		self.populate_b(self.preferences.startup_options)
		scrollbar = tk.Scrollbar(self, orient="vertical")
		scrollbar.config(command=self.listbox_b.yview)
		scrollbar.grid(column=4, row=1,rowspan=4,sticky="ns")
		self.listbox_b.config(yscrollcommand=scrollbar.set)

		#BOTTOM RIGHT PANEL
		frame = tk.Frame(self)
		frame.grid(column=3,row=5,columnspan=2,sticky="news")
		self.split_xlsx_by_ext = tk.BooleanVar(value=self.preferences.startup_split)
		check_button = tk.Checkbutton(frame, text="One extension per Spreadsheet",
									  variable=self.split_xlsx_by_ext,
									  onvalue=True,
									  offvalue=False)
		self.split_jobs = tk.BooleanVar(value=self.preferences.startup_split_jobs)
		check_button.grid(column=0, row=0, sticky="w")
		self.check_buttons.append(check_button)

		check_button = tk.Checkbutton(frame, text="Split gaussian jobs (Link1)",
									  variable=self.split_jobs,
									  onvalue=True,
									  offvalue=False)
		check_button.grid(column=1, row=0, sticky="w")
		self.check_buttons.append(check_button)
		for n in range(2):
			frame.columnconfigure(n,weight=1, uniform='asdffw')

		#CENTER BUTTONS
		button = tk.Button(self, text=">", command=self.move_right, padx="3")
		button.grid(column=2, row=1, sticky="news")
		self.buttons.append(button)
		button = tk.Button(self, text="X", command=self.delete_selection, padx="3")
		button.grid(column=2, row=2, sticky="news")
		self.buttons.append(button)
		button = tk.Button(self, text=u'\u2191', command=self.mv_up_selection, padx="3")
		button.grid(column=2, row=3, sticky="news")
		self.buttons.append(button)
		button = tk.Button(self, text=u'\u2193', command=self.mv_down_selection, padx="3")
		button.grid(column=2, row=4, sticky="news")
		self.buttons.append(button)
		for n in range(4):
			self.rowconfigure(n+1,weight=1, uniform='buttons_')

		#PREFERENCE BUTTONS
		frame = tk.Frame(self)
		frame.grid(column=0,row=6,columnspan=1,rowspan=2)
		top = ["Startup","Preset A","Preset B","Preset C"]
		for i,a in enumerate(top):
			button = tk.Button(frame, text=a, command=lambda a=a: self.load_pref(a))
			button.grid(column=i, row=0,sticky="news",padx="5")
			self.buttons.append(button)
			button = tk.Button(frame, text="Save", command= lambda a=a: self.save_pref(a))
			button.grid(column=i, row=1,sticky="news",padx="5")
			self.buttons.append(button)
		button = tk.Button(frame, text="Add\nAll", command=self.add_all)
		button.grid(column=4, row=0,rowspan=2, sticky="news")
		self.buttons.append(button)
		button = tk.Button(frame, text="Remove\nAll", command=self.rem_all)
		button.grid(column=5, row=0,rowspan=2, sticky="news")
		self.buttons.append(button)
		for n in range(6):
			frame.columnconfigure(n,weight=1, uniform='third')

		#PREVIEW AND GENERATE BUTTONS
		frame = tk.Frame(self)
		frame.grid(column=3,row=6,columnspan=2,rowspan=1,sticky="news")
		button = tk.Button(frame, text="PREVIEW FILES", command=self.preview_files, padx="1")
		button.grid(column=0, row=0, columnspan=1,sticky="news")
		self.buttons.append(button)
		button = tk.Button(frame, text="GENERATE XLS FILE!", command=self.threaded_xls, padx="1")
		button.grid(column=1, row=0, columnspan=1,sticky="news")
		self.buttons.append(button)
		for n in range(2):
			frame.columnconfigure(n,weight=1, uniform='prev')
		#PROGRESS BAR
		self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=100, mode='determinate')
		self.progress["value"] = 0
		self.progress.grid(column=0,row=8,columnspan=6,sticky="news",pady="5")

		#PROGRESS LABEL
		self.progress_label = tk.StringVar()
		self.progress_label.set("github.com/ricalmang")
		label = tk.Label(self, textvariable=self.progress_label)
		label.grid(column=0, row=9,columnspan=6,sticky="e")

		self.path_sheet = ["abs",
		                  "Set A1 to 'abs' if you want Links to use absolute paths",
		                  "Set A1 to anything else if you want Links to use relative paths",
		                  "ON EXCEL: If Hyperlinks are displayed as text and are not working, try:",
		                  "'Ctrl + F' then, replace '=' by '=' in order to force excel to reinterpret cell data",
		                  "ON LIBRE OFFICE: If Hyperlinks are not working, try:",
		                  "'Ctrl + H' (Find & replace) then, replace 'Path!A1' by 'Path.A1' to adjust cell reference of hyperlinks formulas"]
		self.style0 = xlwt.easyxf("", "#.0000000")
	def mv_up_selection(self):
		for n in self.listbox_b.curselection():
			if n == 0: pass
			else:
				text = self.listbox_b.get(n)
				self.listbox_b.delete(n)
				self.listbox_b.insert(n-1,text)
				self.listbox_b.selection_clear(0, tk.END)
				self.listbox_b.selection_set(n - 1)
				self.listbox_b.activate(n - 1)

	def	mv_down_selection(self):
		for n in self.listbox_b.curselection():
			if n == len(self.listbox_b.get(0,tk.END))-1: pass
			else:
				text = self.listbox_b.get(n)
				self.listbox_b.delete(n)
				self.listbox_b.insert(n + 1, text)
				self.listbox_b.selection_clear(0, tk.END)
				self.listbox_b.selection_set(n + 1)
				self.listbox_b.activate(n + 1)

	def load_pref(self,name):
		option = {"Startup" : ["startup_options", "startup_split"  , "startup_split_jobs" ],
				  "Preset A": ["preset_a_options","preset_a_split" , "preset_a_split_jobs"],
				  "Preset B": ["preset_b_options","preset_b_split" , "preset_b_split_jobs"],
				  "Preset C": ["preset_c_options","preset_c_split" , "preset_c_split_jobs"]
				}[name]
		self.populate_b(getattr(self.preferences,option[0])) #UIDS
		self.split_xlsx_by_ext.set(getattr(self.preferences,option[1]))
		self.split_jobs.set(getattr(self.preferences,option[2]))

	def save_pref(self,name):
		uids = " ".join([{a["long"]: a["uid"] for a in self.options}[b] for b in self.listbox_b.get(0,tk.END)])
		result = messagebox.askyesno(title=f"Are you sure?",
									 message=f"This will assign currently selected options to {name} button!",
									 icon='warning')
		if not result: return
		name = {"Startup":"STARTUP","Preset A":"PRESETA","Preset B":"PRESETB","Preset C":"PRESETC"}[name]
		self.preferences.set_variables(name,"options",uids)
		self.preferences.set_variables(name,"splitext", str(self.split_xlsx_by_ext.get()))
		self.preferences.set_variables(name,"splitjobs", str(self.split_jobs.get()))
	def delete_selection(self):
		self.listbox_b.delete(tk.ACTIVE)
	def move_right(self):
		self.listbox_b.insert(tk.END,self.listbox_a.get(tk.ACTIVE))
	def populate_a(self,extension="any"):
		self.listbox_a.delete(0,tk.END)
		if extension=="any":
			for a in self.options:
				self.listbox_a.insert(tk.END, a["long"])
		else:
			for a in self.options:
				if a["extension"] == extension:
					self.listbox_a.insert(tk.END, a["long"])
	def populate_b(self,uids=[]):
		self.listbox_b.delete(0,tk.END)
		for uid in uids:
			self.listbox_b.insert(tk.END, {a["uid"]: a["long"] for a in self.options}[uid])
	def rem_all(self):
		self.listbox_b.delete(0, tk.END)
	def add_all(self):
		self.listbox_b.delete(0, tk.END)
		for a in self.options: self.listbox_b.insert(tk.END, a["long"])

	def evaluate_list(self,folder, recursive=True, extensions=[],errors=set(),files=set(),base_only=False):
		try:
			for file in os.listdir(folder):
				if os.path.isdir(os.path.join(folder, file)):
					if recursive:
						self.evaluate_list(folder=os.path.join(folder, file),
										   recursive=recursive,
										   files=files,
										   extensions=extensions,
										   base_only=base_only,
										   errors=errors)
				elif any([file.endswith(extension) for extension in extensions]):
					files.add(os.path.join(folder,os.path.splitext(file)[0] if base_only else file))
		except PermissionError as error:
			errors.add(error)
		finally:
			return files,errors
	@lock_release
	def preview_files(self):
		global frame_a
		folder = frame_a.in_folder
		cond_1 = folder is None
		cond_2 = type(folder) == str and folder.strip() == ""
		if cond_1 or cond_2:
			messagebox.showinfo(title="Analysis folder is not yet selected!", message="Analysis won't be performed!")
			return
		assert type(folder) == str, type(folder)
		if not os.path.isdir(folder):
			messagebox.showinfo(title="Analysis folder is not a valid path", message="Analysis won't be performed!")
			return
		recursive = frame_a.recursive_analysis.get()
		extensions = list(set(self.dict_options[a][-1] for a in self.listbox_b.get(0,tk.END) if a[-1] != "any"))
		if extensions:
			files,errors = self.evaluate_list(folder=folder, files=set(), recursive=recursive, extensions=extensions)
		else:
			files, errors = [], []
		self.pop_up_files(files,errors)
	def threaded_preview(self):
		job = threading.Thread(target=self.preview_files)
		job.start()
		self.refresh_gui()

	def pop_up_files(self,files,errors):
		top_level = tk.Toplevel()
		top_level.wm_geometry("1000x600")
		scrollbar = tk.Scrollbar(top_level,orient="vertical")
		listbox = tk.Text(top_level,yscrollcommand=scrollbar.set)
		text = "=" * 23 + "THE FOLLOWING FILES WERE FOUND ON THIS DIRECTORY" + "=" * 23 + "\n"
		listbox.insert(tk.INSERT,text)
		text = "\n".join(["{:<4} {}".format(i, trim_str(a, frame_a.str_width)) for i, a in enumerate(files)])
		listbox.insert(tk.INSERT,text)
		if errors:
			text = "\n"*2 + "=" * 10 + "THE FOLLOWING ERRORS WERE RAISED WHILE LOOKING FOR FILES IN THIS DIRECTORY" + "=" * 10 + "\n"
			listbox.insert(tk.INSERT,text)
			text = "\n".join(["{}".format(a) for i, a in enumerate(errors)])
			listbox.insert(tk.INSERT, text)
		listbox.grid(column=0,row=0,sticky="news")
		top_level.grid_columnconfigure(0,weight=1)
		top_level.grid_rowconfigure(0, weight=1)
		scrollbar.config(command=listbox.yview)
		scrollbar.grid(column=1,row=0,sticky="ns")

	def refresh_gui(self):
		self.root.update()
		self.root.after(1000, self.refresh_gui)
	def threaded_xls(self):
		job = threading.Thread(target=self.startup_gen_xls)
		job.start()
		self.refresh_gui()
	@lock_release
	def startup_gen_xls(self):
		global frame_a
		folder = frame_a.in_folder
		# FOLDER MUST BE STRING
		cond_1 = folder is None
		cond_2 = type(folder) == str and folder.strip() == ""
		if cond_1 or cond_2:
			messagebox.showinfo(title="Analysis folder is not yet selected!", message="Analysis won't be performed!")
			return
		need_supl = any([a["supl"] for a in self.options if a["long"] in self.listbox_b.get(0,tk.END)])
		# XLS FOLDER MUST BE PARENT OF IN_FOLDER AND SUPL_FOLDER
		parent = os.path.normpath(os.path.dirname(frame_a.xls_path))
		child_a = os.path.normpath(frame_a.in_folder)
		child_b = os.path.normpath(frame_a.supl_folder)
		cond_1 = not child_a.startswith(parent)
		cond_2 = not child_b.startswith(parent) and need_supl
		if cond_1 and cond_2:
			message = ".xls file must be saved on a folder that contains both suplementary files and input files "
			message += "(May be in subdirectories)."
			message += "\nThis is to ensure that relative path hyperlinks will work properly on the resulting xls file!"
			message += "\nAnalysis won't be performed!"
			messagebox.showinfo(title="Please chose another path scheme!", message=message)
			return
		elif cond_1:
			message = "The folder being analyzed must be contained on a subdirectory of the folder in wich the .xls file is saved."
			message += "\nThis is to ensure that relative path hyperlinks will work properly on the resulting xls file!"
			message += "\nAnalysis won't be performed!"
			messagebox.showinfo(title="Please chose another path scheme!", message=message)
			return
		elif cond_2:
			message  = "The suplementary data folder must be contained on a subdirectory of the folder in wich the .xls file is saved."
			message += "\nThis is to ensure that relative path hyperlinks will work properly on the resulting xls file!"
			message += "\nAnalysis won't be performed!"
			messagebox.showinfo(title="Please chose another path scheme!", message=message)
			return
		# IN_PATH MUST EXIST
		if not os.path.isdir(folder):
			messagebox.showinfo(title="Analysis folder does not exist!",
								message="Analysis won't be performed!")
			return

		# IDEALY SUPL_FOLDER SHOULD NOT EXIST
		if os.path.isdir(child_b) and need_supl:
			message="Do you want to overwrite files on the following directory?\n{}".format(child_b)
			result=messagebox.askyesno(title="Suplementary file directory already exists!", message=message,icon='warning')
			if not result: return

		# IDEALY SUPL_FOLDER BASE DIRECTORY SHOULD EXIST
		if not os.path.isdir(os.path.dirname(child_b)) and need_supl:
			message="Suplementary file directory parent directory does not exists!\nAnalysis won't be performed!\n"
			messagebox.showinfo(title="Parent directory does not exist!", message=message)
			return

		# IDEALY XLS SHOULD NOT EXIST
		if os.path.isfile(frame_a.xls_path):
			message  = "Do you want to overwrite the following file?\n{}".format(frame_a.xls_path)
			message += "\nMoreover, if you want to overwrite it, make sure this file is not opened by another program before you proced."
			result = messagebox.askyesno(title=".xls file already exists!", message=message,icon='warning')
			if not result: return
		# PARENT DIRECTORY XLS SHOULD EXIST
		if not os.path.isdir(os.path.dirname(frame_a.xls_path)):
			messagebox.showinfo(title=".xls parent directory does not exist!",
								message="Analysis won't be performed!")
			return

		recursive = frame_a.recursive_analysis.get()
		extensions = list(set(self.dict_options[a][-1] for a in self.listbox_b.get(0,tk.END) if a[-1] != "any"))

		if extensions:
			basenames, _ = self.evaluate_list(folder=folder, recursive=recursive,files=set(), extensions=extensions, base_only=True)
		else:
			basenames, _ = [], []
		self.progress["value"] = 0
		self.progress_label.set("Analyzing {} file basenames...".format(len(basenames)))
		csv_list = []
		csv_generator = self.analysis_generator(basenames,extensions)
		aux_files_needed = ["LG", "MulkSpinDens", "LastIntCoord", "MulkCharges",
							"ESPCharges", "POPAnalysis", "NPAAnalysis", "APTCharges"]
		for file in csv_generator:
			for request in self.listbox_b.get(0,tk.END):
				key = self.dict_options[request][0]
				if key in aux_files_needed:
					if not file[key] or file[key] == "-": continue
					filename = "_".join([file["Filename"],str(file["Link1"]),key+".txt"])
					filename = os.path.join(frame_a.supl_folder,file["rel_path"],filename)
					try:
						os.makedirs(os.path.dirname(filename), exist_ok=True)
						assert type(file[key]) == str
						with open(filename, "w") as f:
							f.write(file[key])
					except FileExistsError:
						print("Error while creating the following file:")
						print(filename)
						print("File already exists!")
					finally:
						file.update({key: self.mk_hyperlink(filename)})
			csv_list.append(file.copy())
		self.progress["value"] = 100
		self.progress_label.set("Saving xls file...")
		time.sleep(0.1)
		csv_list.sort(key=lambda x: os.path.join(x["rel_path"], x["Filename"]))
		if self.split_xlsx_by_ext.get():
			self.gen_xls_single_ext(csv_list)
		else:
			self.gen_xls_normal(csv_list)
		self.progress_label.set("Done! Please look up: {}".format(trim_str(frame_a.xls_path,frame_a.str_width-50)))
		time.sleep(0.1)
	def gen_xls_single_ext(self,csv_list):
		global frame_a
		wb = Workbook()
		used_extensions = list(dict.fromkeys(self.extension_dict[b] for b in self.used_short_titles))
		x = self.used_short_titles
		data_sheets = []
		if sum(a["Link1"] for a in csv_list) != 0:
			x.insert(0, "Link1")
		for ext in used_extensions:
			if ext =="any":continue
			sheet1 = wb.add_sheet("Data {}".format(ext))
			for i_b, b in enumerate(y for y in x if self.extension_dict[y] in ["any",ext]):
				sheet1.write(0, i_b, b)
			filtered_csv_list = csv_list if ext == ".log" else [a for a in csv_list if a["Link1"]==0]
			data_sheets.append([sheet1,ext,filtered_csv_list])
		sheet2 = wb.add_sheet('Labels')
		for i_b, b in enumerate(x):
			sheet2.write(i_b, 0, b)
			sheet2.write(i_b, 1, self.label_dict[b])
		# TODO Exception: Formula: unknown sheet name Path
		sheet3 = wb.add_sheet('Path')
		for i, a in enumerate(self.path_sheet): sheet3.write(i, 0, a)
		for sheet1,ext,filtered_csv_list in data_sheets:
			for i_a, a in enumerate(filtered_csv_list, start=1):
				for i_b, b in enumerate(y for y in x if self.extension_dict[y] in ["any",ext]):
					args = self.sheet_write_args(i_a,i_b,a,b)
					sheet1.write(*args)

		self.save_xls(wb)
	def sheet_write_args(self,i_a,i_b,a,b):
		if b =="Link1":
			return [i_a, i_b, str(a[b]+1)]
		elif b in self.need_style0 and is_str_float(a[b]):
			return [i_a, i_b, float(a[b]), self.style0]
		elif b in self.need_formula and a[b] not in [None, "-"]:
			return [i_a, i_b, xlwt.Formula(a[b])]
		elif b in self.need_formula and a[b] in [None, "-"]:
			return [i_a, i_b, "-"]
		else:
			return [i_a, i_b, a[b]]

	def gen_xls_normal(self,csv_list):
		global frame_a
		wb = Workbook()
		sheet1 = wb.add_sheet('Data')
		sheet2 = wb.add_sheet('Labels')
		sheet3 = wb.add_sheet('Path')
		for i,a in enumerate(self.path_sheet): sheet3.write(i,0,a)
		x = self.used_short_titles
		if sum(a["Link1"] for a in csv_list) != 0:
			x.insert(0,"Link1")
		for i_b, b in enumerate(x):
			sheet2.write(i_b, 0, b)
			sheet2.write(i_b, 1, self.label_dict[b])
		for i_b, b in enumerate(x):
			sheet1.write(0, i_b, b)
		for i_a, a in enumerate(csv_list, start=1):
			for i_b, b in enumerate(x):
				args = self.sheet_write_args(i_a, i_b, a, b)
				sheet1.write(*args)

		self.save_xls(wb)
	def save_xls(self,wb):
		global frame
		while True:
			try:
				wb.save(frame_a.xls_path)
				break
			except PermissionError:
				result = messagebox.askyesno(title="Error while saving xls file!",
									message="It appears the following file is already open:\n{}\nDo you want to retry to overwrite it?\n(Please close the file before retrying)".format(frame_a.xls_path))
				if not result: break
	def analysis_generator(self,basenames,extensions):
		for i,a in enumerate(basenames):
			self.progress["value"] = int(i / len(basenames) * 100)
			for file_dict in  self.evaluate_file(a,extensions):
				yield file_dict
	def mk_hyperlink(self,x, y="Link"):
		global frame_a
		xls_path = os.path.dirname(frame_a.xls_path)
		return 'HYPERLINK(IF(Path!A1="abs";"{}";"{}");"{}")'.format(x, (os.path.relpath(x, xls_path)), y)
	def evaluate_file(self,a,extensions):
		global frame_a
		x = self.used_short_titles
		row = {"Link1":0}
		row.update({a: "-" for a in x})
		print_exception = lambda e,a: print(f"Error:\n{e}\nOn file:\n{a}")
		up = lambda a: row.update(a)
		file_isfile = lambda name,ext: [y := os.path.normpath(name+ext),os.path.isfile(y)]

		# BLANK
		if (n:="Blank"         ) in x: up({n:" "        })
		if (n:="Eh to kcal/mol") in x: up({n:"627.5095" })
		if (n:="Eh to kJ/mol"  ) in x: up({n:"2625.5002"})

		#FILE PROPERTIES
		up({"Filename":os.path.basename(a)})

		#FOLDER PROPERTIES
		fold_name = os.path.dirname(a)
		up({"rel_path":os.path.relpath(fold_name, frame_a.in_folder)})
		if (n:="Folder") in x: up({n: self.mk_hyperlink(fold_name,row["rel_path"])})

		# XYZ PROPERTIES
		if (ext:=".xyz") in extensions:
			filename, is_file = file_isfile(a,ext)
			up({".xyz":self.mk_hyperlink(filename) if is_file else None})

		#INPUT PROPERTIES .GJF
		if (ext:=".gjf") in extensions:
			filename, is_file = file_isfile(a,ext)
			if (n:= ".gjf" ) in x: up({n: self.mk_hyperlink(filename) if is_file else None})
			if (n:=".gjf_#") in x:
				inp = GjfFile(read_item(filename)) if is_file else False
				up({n:inp.route_text() if inp else "-"})

		#INPUT PROPERTIES .COM
		if (ext:=".com") in extensions:
			filename, is_file = file_isfile(a, ext)
			if (n:=".com") in x: up({n: self.mk_hyperlink(filename) if is_file else None})
			if (n:=".com_#") in x:
				inp = GjfFile(read_item(filename)) if is_file else False
				up({n: inp.route_text() if inp else "-"})

		#INP PROPERTIES
		if (ext:=".inp") in extensions:
			filename, is_file = file_isfile(a, ext)
			if (n:=".inp") in x: up({n: self.mk_hyperlink(filename) if is_file else None})

		#OUT PROPERTIES
		if (ext:=".out") in extensions:
			filename, is_file = file_isfile(a, ext)
			if (n:=".out") in x: up({n: self.mk_hyperlink(filename) if is_file else None})

		#LOG PROPERTIES
		if (ext:=".log") in extensions:
			filename, is_file = file_isfile(a, ext)
			if (n:=".log") in x: up({n: self.mk_hyperlink(filename) if is_file else None})
			other_log_properties = []
			for a in self.listbox_b.get(0,tk.END):
				if self.dict_options[a][-1] == ".log":
					if self.dict_options[a][0] != ".log":
						other_log_properties.append(self.dict_options[a][0])
			if other_log_properties and is_file:
				logs = [LogFile(read_item(filename),self.split_jobs.get()) if is_file else False]
				while True:
					if hasattr(logs[-1],"link_one"): logs.append(getattr(logs[-1],"link_one"))
					else: break
				#print(logs)
				for i,b in enumerate(logs):
					if i > 0:
						yield row
						up({a[1]: "-" for a in other_log_properties})
						up({"Link1":i})
					#try:
					if (n:=".log_#"           )in x: up({n: b.raw_route if b.raw_route else "-"})
					if (n:="E0"               )in x: up({n: b.scf_done[-1][-1] if b.scf_done else "-"})
					if (n:="iFreq"            )in x: up({n: b.last_freq.n_ifreq() if b.last_freq else "-"})
					if (n:="E_ZPE"            )in x: up({n: b.thermal[0] if b.thermal[0] else "-"})
					if (n:="E_tot"            )in x: up({n: b.thermal[1] if b.thermal[1] else "-"})
					if (n:="H_corr"           )in x: up({n: b.thermal[2] if b.thermal[2] else "-"})
					if (n:="G_corr"           )in x: up({n: b.thermal[3] if b.thermal[3] else "-"})
					if (n:="E0+E_ZPE"         )in x: up({n: b.thermal[4] if b.thermal[4] else "-"})
					if (n:="E0+E_tot"         )in x: up({n: b.thermal[5] if b.thermal[5] else "-"})
					if (n:="E0+H_corr"        )in x: up({n: b.thermal[6] if b.thermal[6] else "-"})
					if (n:="E0+G_corr"        )in x: up({n: b.thermal[7] if b.thermal[7] else "-"})
					if (n:="Done?"            )in x: up({n: "Yes" if b.normal_termin else "No"})
					if (n:="Error"            )in x: up({n: b.error_msg if b else "-"})
					if (n:="HOMO"             )in x: up({n: b.homo[-1] if b.homo else "-"})
					if (n:="LUMO"             )in x: up({n: b.lumo[-1] if b.homo else "-"})
					if (n:="HOMO-LUMO"        )in x: up({n: b.homolumo[-1] if b.homolumo else "-"})
					if (n:="Charge"           )in x: up({n: b.charge_mult[0] if b.charge_mult else "-"})
					if (n:="Mult"             )in x: up({n: b.charge_mult[1] if b.charge_mult else "-"})
					if (n:="n_SCF"            )in x: up({n: len(b.scf_done) if b.scf_done else "-"})
					if (n:="n_atoms"          )in x: up({n: b.n_atoms if b.n_atoms else "-"})
					if (n:="TYP"              )in x: up({n: b.calc_type if b.calc_type else "-"})
					if (n:="Needs refinement?")in x: up({n: b.needs_ref()})
					if (n:="S**2 BA"          )in x: up({n: b.s_squared[-1][1] if b.s_squared else "-"})
					if (n:="S**2 After"       )in x: up({n: b.s_squared[-1][2] if b.s_squared else "-"})
					if (n:="LG"               )in x: up({n: "\n".join(b.last_xyz_obj().return_print()) if b.last_xyz_obj() else None})
					if (n:="MulkSpinDens"     )in x: up({n: b.last_muliken_spin_density})
					if (n:="LastIntCoord"     )in x: up({n: b.last_internal_coord })
					if (n:="MulkCharges"      )in x: up({n: b.last_muliken_charges})
					if (n:="ESPCharges"       )in x: up({n: b.last_chelpg_charges })
					if (n:="POPAnalysis"      )in x: up({n: b.pop_analysis        })
					if (n:="NPAAnalysis"      )in x: up({n: b.npa_analysis        })
					if (n:="APTCharges"       )in x: up({n: b.last_apt_charges    })
					#except Exception as e:
					#	print_exception(e,b)
					#finally:
					#	pass
						#print(row)
		yield row
	def _used_short_titles(self):
		return [self.dict_options[a][0] for a in self.listbox_b.get(0, tk.END)]
	used_short_titles = property(_used_short_titles)


#GUI CREATION
root = tk.Tk()
root.title("chemxls v0.0.1")

root_row = 0
root.grid_columnconfigure(0,weight=1)
frame_a = FileFolderSelection(root)
frame_a.grid(column=0,row=root_row,sticky="news",padx="5")
root_row += 1

frame_b = ListBoxFrame(root)
frame_b.grid(column=0,row=root_row,sticky="news",padx="5")
root.grid_rowconfigure(root_row,weight=1)
root_row += 1

w, h = 925 if sys.platform == "win32" or os.name == "nt" else 1000, 685
ws = root.winfo_screenwidth()  # width of the screen
hs = root.winfo_screenheight()  # height of the screen
root.minsize(w, h)
root.maxsize(ws, hs)
x = int(ws / 2 - w / 2)
y = int(hs / 2 - h / 2)
root.geometry(f"{w}x{h}+{x}+{y}")

root.mainloop()
