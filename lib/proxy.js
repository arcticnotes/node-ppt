export class PowerPointObject {

	#ticket;
	#remote; // cache

	constructor( ticket) {
		this.#ticket = ticket;
	}

	get remote() {
		if( this.#remote === undefined)
			this.#remote = this.#ticket.newProxy();
		return this.#remote;
	}
}
