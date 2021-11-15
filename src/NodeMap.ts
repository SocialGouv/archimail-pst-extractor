import long from "long";

/**
 * Stores node names (both alpha and numeric) in node maps for quick lookup.
 */
export class NodeMap {
    private readonly nameToId: Map<string, number> = new Map();

    private readonly idToNumericName: Map<number, long> = new Map();

    private readonly idToStringName: Map<number, string> = new Map();

    /**
     * Set a node into the map.
     */
    public setId(key: unknown, propId: number, idx?: number): void {
        if (typeof key === "number" && idx !== undefined) {
            const lkey = this.transformKey(key, idx);
            this.nameToId.set(lkey.toString(), propId);
            this.idToNumericName.set(propId, lkey);
            // console.log('NodeMap::setId: propId = ' + propId + ', lkey = ' + lkey.toString());
        } else if (typeof key === "string") {
            this.nameToId.set(key, propId);
            this.idToStringName.set(propId, key);
            // console.log('NodeMap::setId: propId = ' + propId + ', key = ' + key);
        } else {
            throw new Error(`NodeMap::setId bad param type ${typeof key}`);
        }
    }

    /**
     * Get a node from the map.
     */
    public getId(key: unknown, idx?: number): number {
        let id: number | undefined = undefined;
        if (typeof key === "number" && idx) {
            id = this.nameToId.get(this.transformKey(key, idx).toString());
        } else if (typeof key === "string") {
            id = this.nameToId.get(key);
        } else {
            throw new Error(`NodeMap::getId bad param type ${typeof key}`);
        }
        if (!id) {
            return -1;
        }
        return id;
    }

    /**
     * Get a node from the map.
     */
    public getNumericName(propId: number): long | undefined {
        return this.idToNumericName.get(propId);
    }

    private transformKey(key: number, idx: number): long {
        let lidx = long.fromNumber(idx);
        lidx = lidx.shiftLeft(32);
        lidx = lidx.or(key);
        return lidx;
    }
}
