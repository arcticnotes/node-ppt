import { Collection} from '@arcticnotes/node-wsh';
import { PowerPointObject} from './proxy.js';
import { Slide} from './Slide.js';

export class Presentation extends PowerPointObject {

	#ticket;

	constructor( ticket) {
		super( ticket);
		this.#ticket = ticket;
	}

	get name() {
		return this.#ticket.get( 'Name');
	}

	get path() {
		return this.#ticket.get( 'Path');
	}

	get width() {
		return this.remote.PageSetup.SlideWidth;
	}

	set width( width) {
		this.remote.PageSetup.SlideWidth = width
	}

	get height() {
		return this.remote.PageSetup.SlideHeight;
	}

	set height( height) {
		this.remote.PageSetup.SlideHeight = height
	}

	get layouts() {
		return this.#ticket.get( 'SlideMaster', ticket => new SlideMaster( ticket)).layouts;
	}

	get slides() {
		return this.#ticket.get( 'Slides', ticket => new Slides( ticket));
	}

	newSlide( layout, index = undefined) {
		const indexBase1 = index !== undefined? index + 1: this.slides.length + 1;
		return this.slides[ ADD_SLIDE]( indexBase1, layout);
	}
}

const ADD_SLIDE = Symbol();

class Slides extends Collection {

	#ticket;

	constructor( ticket) {
		super( ticket, Slide);
		this.#ticket = ticket;
	}

	[ ADD_SLIDE]( indexBase1, layout) {
		return this.#ticket.get( 'AddSlide', ticket => ( ...args) => ticket.apply( this, args, t => new Slide( t)))(
				indexBase1, layout);
	}
}

class SlideMaster extends PowerPointObject {

	#ticket;

	constructor( ticket) {
		super( ticket);
		this.#ticket = ticket;
	}

	get layouts() {
		return this.#ticket.get( 'CustomLayouts', ticket => new Collection( ticket, Layout));
	}
}

export class Layout extends PowerPointObject {

	#ticket;

	constructor( ticket) {
		super( ticket);
		this.#ticket = ticket;
	}

	get name() {
		return this.#ticket.get( 'Name');
	}
}
