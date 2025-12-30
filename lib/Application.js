import { Collection, WindowsScriptingHost} from '@arcticnotes/node-wsh';
import { PowerPointObject} from './proxy.js';
import { Presentation} from './Presentation.js';

export class Application extends PowerPointObject {

	/**
	 * Finds a running PowerPoint.Application instance and communicate with it through Windows Scripting Host. If no such
	 * program is running, an Error is thrown. To shutdown cleanly, call {@link Application.detach()} after use.
	 *
	 * @param { object} options Options. Set `node-wsh` options with `node-wsh` property.
	 */
	static async attach( options = {}) {
		const wsh = await WindowsScriptingHost.connect( options[ 'node-wsh']);
		const GetObject = wsh.global( 'GetObject', ticket => function( path, progId) {
			if( path === '' && progId === 'PowerPoint.Application')
				return ticket.apply( undefined, [ path, progId], appTicket => new Application( appTicket, wsh));
		});
		return GetObject( '', 'PowerPoint.Application');
	}

	#ticket;
	#wsh;

	constructor( ticket, wsh) {
		super( ticket);
		this.#ticket = ticket;
		this.#wsh = wsh;
	}

	async detach() {
		await this.#wsh.disconnect();
	}

	get presentations() {
		return this.#ticket.get( 'Presentations', ticket => new Collection( ticket, Presentation));
	}
}
