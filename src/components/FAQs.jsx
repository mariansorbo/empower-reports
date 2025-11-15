import React from 'react'
import './FAQs.css'

const FAQs = ({ onBack }) => {
  return (
    <div className="faqs-page">
      <div className="faqs-container">
        <header className="faqs-header">
          <button className="faqs-back-btn" onClick={onBack}>
            ← Back to home
          </button>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px', marginBottom: '10px' }}>
            <span style={{ fontSize: '32px' }}>📊</span>
            <span style={{ fontSize: '20px', fontWeight: '600', color: '#333' }}>
              Empower <span style={{ color: '#667eea' }}>Reports</span>
            </span>
          </div>
          <h1>❓ Frequently Asked Questions (FAQs)</h1>
        </header>

        <div className="faqs-content">
          <section className="faq-item">
            <div className="faq-icon">📊</div>
            <div className="faq-content">
              <h2>What exactly does Report Tuner do?</h2>
              <p>
                Report Tuner converts your Power BI .pbit files into interactive technical documentation, 
                where you can navigate all tables, metrics, relationships and internal dependencies. 
                Ideal for understanding, auditing or scaling complex models.
              </p>
            </div>
          </section>

          <section className="faq-item">
            <div className="faq-icon">🧠</div>
            <div className="faq-content">
              <h2>Do I need technical knowledge to use it?</h2>
              <p>
                No. Report Tuner was designed for business, data and development teams. 
                The interface is simple, and the result is clear and visual. If you know how to use Power BI, 
                you can understand Report Tuner.
              </p>
            </div>
          </section>

          <section className="faq-item">
            <div className="faq-icon">🧾</div>
            <div className="faq-content">
              <h2>What kind of information does it return?</h2>
              <p>The analysis includes:</p>
              <ul>
                <li>List of tables and columns used</li>
                <li>Relationships between tables</li>
                <li>Dependencies between measures</li>
                <li>DAX formulas used</li>
                <li>Unused fields (optional)</li>
                <li>Downloadable files</li>
              </ul>
            </div>
          </section>

          <section className="faq-item">
            <div className="faq-icon">🧪</div>
            <div className="faq-content">
              <h2>Is this a free trial?</h2>
              <p>
                Yes. You are using a free trial version. All analyses you do at this stage 
                are free, and your feedback helps us improve the tool.
              </p>
            </div>
          </section>

          <section className="faq-item">
            <div className="faq-icon">📁</div>
            <div className="faq-content">
              <h2>What files can I upload?</h2>
              <p>
                We accept Power BI .pbit files. Soon we will enable compatibility with correctly 
                exported .pbix files.
              </p>
            </div>
          </section>

          <section className="faq-item">
            <div className="faq-icon">🛠️</div>
            <div className="faq-content">
              <h2>What if I have errors or need support?</h2>
              <p>
                You can contact us from the bottom section of the page. We are testing the platform 
                and your feedback is key.
              </p>
            </div>
          </section>

          <section className="faq-item faq-item-highlight">
            <div className="faq-icon">🔒</div>
            <div className="faq-content">
              <h2>What happens to my files? Are they private? Are they sold?</h2>
              <p>
                Yes, your .pbit files are processed completely securely and are never sold or shared 
                with third parties.
              </p>
              <p>
                We do not store your data beyond the time strictly necessary to process and return 
                the result. The only use we give them is to generate the technical analysis you requested.
              </p>
              <p>
                We also do not use your files to train artificial intelligence models, nor do we reuse them 
                for any purpose other than what you requested.
              </p>
              <p className="faq-emphasis">
                Your file is yours, and we only use it to deliver the final product.
              </p>
            </div>
          </section>

          <section className="faq-item">
            <div className="faq-icon">📁</div>
            <div className="faq-content">
              <h2>Why does using .pbit files improve security?</h2>
              <p>
                The .pbit format (Power BI Template) does not contain sensitive or production data, only structure 
                of the model, relationships, DAX measures and metadata.
              </p>
              <p>This means that:</p>
              <ul>
                <li>It does not include raw data from your sources.</li>
                <li>It does not expose confidential information about customers, revenue or metrics.</li>
                <li>It's safe to share between teams without compromising privacy.</li>
              </ul>
              <p>
                Using .pbit is a good practice for auditing, versioning and documenting your reports without risks.
              </p>
            </div>
          </section>

          <section className="faq-item faq-item-highlight">
            <div className="faq-icon">🔐</div>
            <div className="faq-content">
              <h2>How do you protect the confidentiality of the analysis?</h2>
              <ul>
                <li>We use encrypted connections (HTTPS) at all times.</li>
                <li>
                  The file is automatically deleted once the analysis is finished (or when the user decides).
                </li>
                <li>
                  If you choose to save the results, only you have access through your account.
                </li>
                <li>By contacting us.</li>
              </ul>
              <p>
                Report Tuner was designed from the beginning as a technical tool focused on security, 
                ideal for teams working with sensitive information but needing documentation and collaboration.
              </p>
            </div>
          </section>
        </div>
      </div>
    </div>
  )
}

export default FAQs
