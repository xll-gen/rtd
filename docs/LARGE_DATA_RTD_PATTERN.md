# Architectural Advice: Handling Large Matrix Data with RTD and Go

## Executive Summary

**Question:** Is a parameter hash needed when receiving large data like matrices?
**Answer:** **Yes.** A parameter hash (or a content-based unique identifier) is highly recommended and effectively necessary for this architecture.

## The Challenge

In a standard Excel RTD (Real-Time Data) implementation, the `ConnectData` method receives an array of strings (`Topic` strings) to identify the data stream.

```cpp
HRESULT ConnectData(long TopicID, SAFEARRAY** Strings, bool* GetNewValues, VARIANT* pvarOut);
```

When dealing with large input parameters (like a 100x100 matrix) that determine the result:
1.  **Size Limit:** You cannot pass a large matrix as a set of topic strings.
2.  **Referential Transparency:** If the matrix content changes, the result should change. This means the RTD "Topic" must change to trigger a new subscription.
3.  **IPC Lookup:** The external Go server needs a key to look up the matrix data.

## Recommended Architecture: "Hash-as-Topic"

To solve this, we use a **Hybrid XLL/RTD** approach where the XLL function handles the data hand-off and the RTD server handles the subscription.

### Data Flow

1.  **Excel Calculation (XLL Side):**
    *   The user calls an XLL function, e.g., `=MyService.ProcessMatrix(A1:Z100)`.
    *   The XLL function receives the matrix data (e.g., via `LPXLOPER12`).
    *   **Step A (Hash):** The XLL calculates a deterministic **Hash** of the matrix content (e.g., using xxHash or SHA-256).
    *   **Step B (Send):** The XLL checks if this Hash is already known/sent to the Go server. If not, it sends the pair `{Hash, MatrixData}` to the Go server via IPC (e.g., gRPC, ZeroMQ, Shared Memory).
    *   **Step C (Subscribe):** The XLL function returns a call to the RTD function using the Hash as the topic:
        *   `return RTD("My.ProgID", "", "matrix_result", "<HASH_STRING>");`

2.  **RTD Connection (C++ Server):**
    *   Excel calls `ConnectData`. The `Strings` array contains `"matrix_result"` and `"<HASH_STRING>"`.
    *   The RTD Server extracts the Hash.
    *   It sends a **Subscription Request** for this Hash to the Go server.

3.  **Processing (Go Server):**
    *   The Go server receives the Subscription Request.
    *   It looks up the Matrix Data using the Hash (received in Step B).
        *   *Note:* If the subscription arrives before the data (race condition), the Go server should wait/buffer the request.
    *   The Go server processes the data and streams updates back to the C++ RTD Server.

4.  **Update (Excel):**
    *   The RTD Server receives the update, calls `UpdateNotify`, and returns the result via `RefreshData`.

## Why is the Hash Needed?

### 1. Unique Topic Identification
Excel's dependency engine needs to know when to un-subscribe from an old stream and subscribe to a new one.
*   **With Hash:** If cell `A1` changes from `1` to `2`, the Matrix Hash changes. The XLL returns a *new* RTD string. Excel disconnects the old Topic and connects the new one.
*   **Without Hash (e.g., Random ID):** A new ID is generated on every recalculation (even if data is same), causing unnecessary reconnects/recomputations.
*   **Without Hash (e.g., Static ID):** If you use a static ID (like Cell Address), Excel won't know the input data changed, and the RTD server won't receive a new `ConnectData` call.

### 2. Data Deduplication
If two different cells (or sheets) pass the exact same matrix:
*   They generate the **same Hash**.
*   Excel (often) optimizes this to a single RTD Topic connection.
*   Even if Excel creates two Topic IDs, they both subscribe to the same Hash on the Go server.
*   The Go server computes the result **once** and broadcasts it to all subscribers.

## Implementation Advice

### Hashing Algorithm
Use a fast, non-cryptographic hash function if security is not a concern (internal data).
*   **xxHash (XXH3_64bits or XXH128):** Extremely fast, suitable for large memory buffers.
*   **MurmurHash3:** Common alternative.
*   **SHA-256:** Use if cryptographic collision resistance is strictly required (slower).

### Handling the "Race Condition"
Since the XLL "Send Data" and RTD "Connect" are asynchronous events:
*   **Scenario:** XLL sends data to Go, but RTD connects before Go processes the data packet.
*   **Solution:** The Go server should maintain a "Pending Subscriptions" map. If a subscription comes for a Hash that isn't found yet, hold it. When the Data packet arrives, match it and trigger the processing.

### C++ Snippet (Conceptual)

```cpp
// Inside XLL Function
extern "C" LPXLOPER12 WINAPI ProcessMatrix(LPXLOPER12 pMatrix) {
    // 1. Hash the Matrix
    std::string hash = ComputeHash(pMatrix);

    // 2. Send to Go (Async/Non-blocking recommended)
    // If we haven't sent this hash recently...
    if (!Cache.Has(hash)) {
        GoClient.SendData(hash, pMatrix);
        Cache.MarkSent(hash);
    }

    // 3. Return RTD Call
    // Wrapper around xlCall(xlfRtd, ...)
    return CallRtd("MyServer", "MatrixTopic", hash);
}
```

## Summary
Using a **content-based hash** solves the problem of bridging the gap between Excel's cell-based dependency system and your external Go calculation server. It provides a robust, efficient, and "Excel-friendly" way to handle large datasets.
