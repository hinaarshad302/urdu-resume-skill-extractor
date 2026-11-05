                    for item in experience:
                        st.write("• " + item)
                else:
                    st.write("No experience information detected.")

            with st.expander("📄 View Original Extracted Resume Text"):
                st.text_area("Extracted document content", resume_text, height=250)

            # -----------------------------------------
            # EDITABLE PROFILE + MULTIFORMAT EXPORT
            # -----------------------------------------
            st.markdown('<div class="section-heading">07 &nbsp; Edit Candidate Profile</div>', unsafe_allow_html=True)
            st.write("Review and correct the extracted information before creating a CV.")

            with st.form("candidate_profile_form"):
                col_a, col_b = st.columns(2)
                with col_a:
                    candidate_name = st.text_input("Candidate Name / امیدوار کا نام", value=name)
                    edited_email = st.text_input("Email / ای میل", value=email)
                with col_b:
                    edited_phone = st.text_input("Phone / فون نمبر", value=phone)
                    cv_language = st.selectbox(
                        "CV Language / CV کی زبان",
                        ["English", "Urdu", "Bilingual"]
                    )

                summary = st.text_area(
                    "Professional Summary / پیشہ ورانہ تعارف",
                    placeholder="Write a short professional introduction..."
                )
                edited_skills = st.text_area(
                    "Skills / مہارتیں",
                    value=", ".join(skills),
                    height=90
                )
                edited_education = st.text_area(
                    "Education / تعلیم",
                    value="\n".join(education),
                    height=100
                )
                edited_experience = st.text_area(
                    "Experience / تجربہ",
                    value="\n".join(experience),
                    height=100
                )
                submitted = st.form_submit_button("Prepare CV Downloads")

            if submitted:
                st.session_state["candidate_profile"] = {
                    "name": candidate_name.strip(),
                    "email": edited_email.strip(),
                    "phone": edited_phone.strip(),
                    "summary": summary.strip(),
                    "skills": edited_skills.strip(),
                    "education": edited_education.strip(),
                    "experience": edited_experience.strip()
                }
                st.session_state["cv_language"] = cv_language

            profile = st.session_state.get("candidate_profile")
            language = st.session_state.get("cv_language", "English")

            if profile:
                st.markdown('<div class="section-heading">08 &nbsp; Download Professional CV</div>', unsafe_allow_html=True)
                st.caption(f"Selected CV language: {language}")

                d1, d2, d3 = st.columns(3)
                with d1:
                    st.download_button(
                        "⬇ Word CV (.docx)",
                        data=build_docx(profile, language),
                        file_name="professional_cv.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                with d2:
                    st.download_button(
                        "⬇ Excel Profile (.xlsx)",
                        data=build_excel(profile),
                        file_name="candidate_profile.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with d3:
                    try:
                        pdf_data = build_pdf(profile, language)
                        st.download_button(
                            "⬇ PDF CV (.pdf)",
                            data=pdf_data,
                            file_name="professional_cv.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    except Exception as pdf_error:
                        st.warning(str(pdf_error))

                st.info(
                    "Important: extracted information is rule-based. Please verify "
                    "all details. The app does not automatically translate your "
                    "personal information; enter Urdu content yourself for an Urdu CV."
                )

    except Exception as error:
        st.error("Unable to process this resume.")
        st.exception(error)
else:
    st.info("Upload a resume above to start candidate analysis.")


# =====================================================
# FOOTER
# =====================================================
st.markdown("""
<div class="footer">
    <b>Urdu Resume Skill Extractor</b><br>
    Academic Final Year Project | 2026–2027<br>
    Python • Streamlit • Document Information Extraction<br><br>
    Prototype output should be reviewed before any recruitment decision.
</div>
""", unsafe_allow_html=True)
