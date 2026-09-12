/*
 * BundleAccessor: EventManager の Accessor 互換。
 * 個別ファイルを fetch する代わりに、1 度読み込んだ bundle.json から返す。
 *
 * bundle = { eventId, files: { "<相対パス>": <パース済みJSON>, ... } }
 * 例: "Event.json", "Pilots.json", "Rounds.json", "Stages.json",
 *     "httpfiles/Channels.json", "<raceId>/Race.json", "<raceId>/Result.json"
 *
 * EventManager は eventDirectory を "" で生成する前提なので、
 *   GetEvent -> "/Event.json"
 *   GetRace  -> "/<id>/Race.json"
 *   GetChannels -> "httpfiles/Channels.json"
 * のようなキーになる。先頭の "./" や "/" は正規化して引く。
 */
class BundleAccessor {
    constructor(bundle) {
        this.files = (bundle && bundle.files) || {};
    }

    _key(url) {
        let k = String(url == null ? "" : url);
        k = k.replace(/^\.\//, "");
        k = k.replace(/^\/+/, "");
        return k;
    }

    async GetJSON(url) {
        const v = this.files[this._key(url)];
        return v == null ? [] : v;
    }
}
