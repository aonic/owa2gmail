"""
POP e-mail server for Microsoft Outlook Web Access scraper

This wraps the Outlook Web Access scraper by providing a POP interface to it.

Run this file from the command line to start the server on localhost port 8110.
The server will run until you end the process. If you specify the "--once"
command-line option, it'll only run one full POP transaction and quit. The
latter is useful for the "precommand" option in KMail, which you can configure
to run this script every time you check e-mail so it starts the server in the
background.
"""

# Forked from:
#   Adrian Holovaty's weboutlook project
#
# By:
#   Raja Kapur <raja.kapur@gmail.com>
#
# https://github.com/aonic/owa2gmail
# weboutlook: http://code.google.com/p/weboutlook/

# Based on gmailpopd.py by follower@myrealbox.com,
# which was in turn based on smtpd.py by Barry Warsaw.
#
# Copyright (C) 2006 Adrian Holovaty <holovaty@gmail.com>
#
# This program is free software; you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free Software
# Foundation; either version 2 of the License, or (at your option) any later
# version.
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details.
#
# You should have received a copy of the GNU General Public License along with
# this program; if not, write to the Free Software Foundation, Inc., 59 Temple
# Place, Suite 330, Boston, MA 02111-1307 USA

import asyncore, asynchat, socket, sys

# relative import
from scraper import InvalidLogin, OutlookWebScraper

__version__ = 'Python Outlook Web Access POP3 proxy version 0.0.2'

TERMINATOR = '\r\n'

def quote_dots(lines):
	for line in lines:
		if line.startswith("."):
			line = "." + line
		yield line

class POPChannel(asynchat.async_chat):
	def __init__(self, conn, options):
		self.webmail_server = options.webmail_server
		self.unread_messages = options.unread is True
		asynchat.async_chat.__init__(self, conn)
		self.__line = []
		self.push('+OK %s %s' % (socket.getfqdn(), __version__))
		self.set_terminator(TERMINATOR)
		self._activeDataChannel = None

	# Overrides base class for convenience
	def push(self, msg):
		#print msg
		asynchat.async_chat.push(self, msg + TERMINATOR)

	# Implementation of base class abstract method
	def collect_incoming_data(self, data):
		self.__line.append(data)

	# Implementation of base class abstract method
	def found_terminator(self):
		line = ''.join(self.__line)
		self.__line = []
		if not line:
			self.push('500 Error: bad syntax')
			return
		method = None
		i = line.find(' ')
		if i < 0:
			command = line.upper()
			arg = None
		else:
			command = line[:i].upper()
			arg = line[i+1:].strip()
		method = getattr(self, 'pop_' + command, None)
		if not method:
			self.push('-ERR Error : command "%s" not implemented' % command)
			return
		method(arg)
		return

	def pop_USER(self, user):
		# Logs in any username.
		if not user:
			self.push('-ERR: Syntax: USER username')
		else:
			self.username = user # Store for later.
			self.push('+OK Password required')

	def pop_PASS(self, password=''):
		self.scraper = OutlookWebScraper(self.webmail_server, self.username, password)
		try:
			self.scraper.login()
		except InvalidLogin:
			self.push('-ERR Login failed. (Wrong username/password?)')
		else:
			self.push('+OK User logged in')
			self.inbox_cache = self.scraper.inbox(self.unread_messages)
			self.msg_cache = [self.scraper.get_message(msg_id) for msg_id in self.inbox_cache]

	def pop_STAT(self, arg):
		dropbox_size = sum([len(msg) for msg in self.msg_cache])
		self.push('+OK %d %d' % (len(self.inbox_cache), dropbox_size))

	def pop_UIDL(self, arg):
		if not arg:
			self.push('+OK')
			for i, msg in enumerate(self.msg_cache):
				self.push('%d %d' % (i+1, i+1))
			self.push(".")
		else:
			self.push('+OK %s %s' % (arg, arg))

	def pop_LIST(self, arg):
		if not arg:
			self.push('+OK')
			for i, msg in enumerate(self.msg_cache):
				self.push('%d %d' % (i+1, len(msg)))
			self.push(".")
		else:
			# TODO: Handle per-msg LIST commands
			raise NotImplementedError

	def pop_RETR(self, arg):
		if not arg:
			self.push('-ERR: Syntax: RETR msg')
		else:
			# TODO: Check request is in range.
			msg_index = int(arg) - 1
			msg_id = self.inbox_cache[msg_index]
			msg = self.msg_cache[msg_index]
			msg = msg.lstrip() + TERMINATOR

			self.push('+OK')

			for line in quote_dots(msg.split(TERMINATOR)):
				self.push(line)
			self.push('.')

			# Delete the message
			self.scraper.delete_message(msg_id)

	def pop_TOP(self, arg):
		if not arg:
			self.push('-ERR: Syntax: TOP msg')
		else:
			# TODO: Check request is in range.
			msg_index = int(arg.split(' ')[0]) - 1
			msg = self.msg_cache[msg_index]
			msg = msg.split('\r\n\r\n')[0]
			msg = msg.lstrip() + TERMINATOR

			self.push('+OK')

			for line in quote_dots(msg.split(TERMINATOR)):
				self.push(line)
			self.push('.')

	def pop_QUIT(self, arg):
		self.push('+OK Goodbye')
		self.close_when_done()

	def handle_error(self):
		asynchat.async_chat.handle_error(self)

class POP3Proxy(asyncore.dispatcher):
	def __init__(self, localaddr, options):
		"""
		localaddr is a tuple of (ip_address, port).

		options.unread is a boolean specifying whether the server should only retrieve unread message.
		
		options.webmail_server is the subdomain where OWA is accessed
		"""
		self.options = options
		asyncore.dispatcher.__init__(self)
		self.create_socket(socket.AF_INET, socket.SOCK_STREAM)
		# try to re-use a server port if possible
		self.set_reuse_addr()
		self.bind(localaddr)
		self.listen(5)

	def handle_accept(self):
		conn, addr = self.accept()
		channel = POPChannel(conn, self.options)

if __name__ == '__main__':
	from optparse import OptionParser
	parser = OptionParser("usage: %prog [options]")
	parser.add_option('--unread', action='store_true', dest='unread',
		help='Only get unread messages (gets all messages from first page by default.)')
	parser.add_option('--owa-server', dest='webmail_server',
		help='Required: URL to Outlook Web Access server. (eg: https://webmail.example.com)')
	options, args = parser.parse_args()

	if(options.webmail_server == None):
		print '%s: %s' % (sys.argv[0], "--owa-server is a required option")
        print '%s: Try --help for usage details.' % (sys.argv[0])
        sys.exit(0)

	proxy = POP3Proxy(('0.0.0.0', 2221), options)
	try:
		asyncore.loop()
	except KeyboardInterrupt:
		pass
